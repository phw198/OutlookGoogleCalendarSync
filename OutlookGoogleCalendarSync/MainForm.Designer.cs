namespace OutlookGoogleCalendarSync
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
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
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabSettings = new System.Windows.Forms.TabControl();
            this.tabPage_Sync = new System.Windows.Forms.TabPage();
            this.cbVerboseOutput = new System.Windows.Forms.CheckBox();
            this.lNextSyncVal = new System.Windows.Forms.Label();
            this.lLastSyncVal = new System.Windows.Forms.Label();
            this.lNextSync = new System.Windows.Forms.Label();
            this.lLastSync = new System.Windows.Forms.Label();
            this.LogBox = new System.Windows.Forms.TextBox();
            this.bSyncNow = new System.Windows.Forms.Button();
            this.tabPage_Settings = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbHelp = new System.Windows.Forms.TextBox();
            this.gbGoogle = new System.Windows.Forms.GroupBox();
            this.lGoogleHelp = new System.Windows.Forms.Label();
            this.lGoogleCalendar = new System.Windows.Forms.Label();
            this.bGetGoogleCalendars = new System.Windows.Forms.Button();
            this.cbGoogleCalendars = new System.Windows.Forms.ComboBox();
            this.gbOutlook = new System.Windows.Forms.GroupBox();
            this.lOutlookCalendar = new System.Windows.Forms.Label();
            this.cbOutlookCalendars = new System.Windows.Forms.ComboBox();
            this.rbOutlookEWS = new System.Windows.Forms.RadioButton();
            this.rbOutlookDefaultMB = new System.Windows.Forms.RadioButton();
            this.rbOutlookAltMB = new System.Windows.Forms.RadioButton();
            this.gbEWS = new System.Windows.Forms.GroupBox();
            this.txtEWSServerURL = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtEWSPass = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtEWSUser = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.ddMailboxName = new System.Windows.Forms.ComboBox();
            this.gbAppBehaviour = new System.Windows.Forms.GroupBox();
            this.cbStartOnStartup = new System.Windows.Forms.CheckBox();
            this.cbShowBubbleTooltips = new System.Windows.Forms.CheckBox();
            this.cbMinimizeToTray = new System.Windows.Forms.CheckBox();
            this.cbStartInTray = new System.Windows.Forms.CheckBox();
            this.cbCreateFiles = new System.Windows.Forms.CheckBox();
            this.bSave = new System.Windows.Forms.Button();
            this.gbSyncOptions = new System.Windows.Forms.GroupBox();
            this.cbOutlookPush = new System.Windows.Forms.CheckBox();
            this.cbMergeItems = new System.Windows.Forms.CheckBox();
            this.syncDirection = new System.Windows.Forms.ComboBox();
            this.cbDisableDeletion = new System.Windows.Forms.CheckBox();
            this.lMiscOptions = new System.Windows.Forms.Label();
            this.cbConfirmOnDelete = new System.Windows.Forms.CheckBox();
            this.cbAddReminders = new System.Windows.Forms.CheckBox();
            this.lAttributes = new System.Windows.Forms.Label();
            this.cbAddAttendees = new System.Windows.Forms.CheckBox();
            this.cbIntervalUnit = new System.Windows.Forms.ComboBox();
            this.cbAddDescription = new System.Windows.Forms.CheckBox();
            this.tbInterval = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.tbDaysInTheFuture = new System.Windows.Forms.NumericUpDown();
            this.tbDaysInThePast = new System.Windows.Forms.NumericUpDown();
            this.lDaysInFuture = new System.Windows.Forms.Label();
            this.lDaysInPast = new System.Windows.Forms.Label();
            this.lDateRange = new System.Windows.Forms.Label();
            this.tabPage_About = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.pbDonate = new System.Windows.Forms.PictureBox();
            this.lAboutURL = new System.Windows.Forms.LinkLabel();
            this.lAboutMain = new System.Windows.Forms.Label();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.tabSettings.SuspendLayout();
            this.tabPage_Sync.SuspendLayout();
            this.tabPage_Settings.SuspendLayout();
            this.panel1.SuspendLayout();
            this.gbGoogle.SuspendLayout();
            this.gbOutlook.SuspendLayout();
            this.gbEWS.SuspendLayout();
            this.gbAppBehaviour.SuspendLayout();
            this.gbSyncOptions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInTheFuture)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInThePast)).BeginInit();
            this.tabPage_About.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).BeginInit();
            this.SuspendLayout();
            // 
            // tabSettings
            // 
            this.tabSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabSettings.Controls.Add(this.tabPage_Sync);
            this.tabSettings.Controls.Add(this.tabPage_Settings);
            this.tabSettings.Controls.Add(this.tabPage_About);
            this.tabSettings.Location = new System.Drawing.Point(12, 12);
            this.tabSettings.Name = "tabSettings";
            this.tabSettings.SelectedIndex = 0;
            this.tabSettings.Size = new System.Drawing.Size(495, 568);
            this.tabSettings.TabIndex = 0;
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
            this.bSyncNow.Text = "Start Sync";
            this.bSyncNow.UseVisualStyleBackColor = true;
            this.bSyncNow.Click += new System.EventHandler(this.Sync_Click);
            // 
            // tabPage_Settings
            // 
            this.tabPage_Settings.Controls.Add(this.panel1);
            this.tabPage_Settings.Controls.Add(this.gbGoogle);
            this.tabPage_Settings.Controls.Add(this.gbOutlook);
            this.tabPage_Settings.Controls.Add(this.gbAppBehaviour);
            this.tabPage_Settings.Controls.Add(this.bSave);
            this.tabPage_Settings.Controls.Add(this.gbSyncOptions);
            this.tabPage_Settings.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Settings.Name = "tabPage_Settings";
            this.tabPage_Settings.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Settings.Size = new System.Drawing.Size(487, 542);
            this.tabPage_Settings.TabIndex = 1;
            this.tabPage_Settings.Text = "Settings";
            this.tabPage_Settings.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tbHelp);
            this.panel1.Location = new System.Drawing.Point(218, 375);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(263, 90);
            this.panel1.TabIndex = 19;
            // 
            // tbHelp
            // 
            this.tbHelp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbHelp.BackColor = System.Drawing.SystemColors.Window;
            this.tbHelp.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbHelp.Cursor = System.Windows.Forms.Cursors.Help;
            this.tbHelp.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.tbHelp.Location = new System.Drawing.Point(9, 21);
            this.tbHelp.Margin = new System.Windows.Forms.Padding(10);
            this.tbHelp.Multiline = true;
            this.tbHelp.Name = "tbHelp";
            this.tbHelp.Size = new System.Drawing.Size(244, 54);
            this.tbHelp.TabIndex = 18;
            this.tbHelp.Text = "It\'s advisable to create a dedicated calendar in Google for synchronising to from" +
                " Outlook. Otherwise you may end up with duplicates or non-Outlook entries delete" +
                "d.";
            this.tbHelp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // gbGoogle
            // 
            this.gbGoogle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbGoogle.Controls.Add(this.lGoogleHelp);
            this.gbGoogle.Controls.Add(this.lGoogleCalendar);
            this.gbGoogle.Controls.Add(this.bGetGoogleCalendars);
            this.gbGoogle.Controls.Add(this.cbGoogleCalendars);
            this.gbGoogle.Location = new System.Drawing.Point(6, 157);
            this.gbGoogle.Name = "gbGoogle";
            this.gbGoogle.Size = new System.Drawing.Size(475, 75);
            this.gbGoogle.TabIndex = 5;
            this.gbGoogle.TabStop = false;
            this.gbGoogle.Text = "Google";
            // 
            // lGoogleHelp
            // 
            this.lGoogleHelp.AutoSize = true;
            this.lGoogleHelp.Location = new System.Drawing.Point(157, 13);
            this.lGoogleHelp.Name = "lGoogleHelp";
            this.lGoogleHelp.Size = new System.Drawing.Size(311, 26);
            this.lGoogleHelp.TabIndex = 4;
            this.lGoogleHelp.Text = "If this is the first time, you\'ll need to authorise the app to connect.\r\nDoesn\'t " +
                "take long - just follow the steps :)";
            // 
            // lGoogleCalendar
            // 
            this.lGoogleCalendar.Location = new System.Drawing.Point(9, 45);
            this.lGoogleCalendar.Name = "lGoogleCalendar";
            this.lGoogleCalendar.Size = new System.Drawing.Size(81, 14);
            this.lGoogleCalendar.TabIndex = 3;
            this.lGoogleCalendar.Text = "Select calendar";
            // 
            // bGetGoogleCalendars
            // 
            this.bGetGoogleCalendars.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetGoogleCalendars.Location = new System.Drawing.Point(12, 16);
            this.bGetGoogleCalendars.Name = "bGetGoogleCalendars";
            this.bGetGoogleCalendars.Size = new System.Drawing.Size(139, 20);
            this.bGetGoogleCalendars.TabIndex = 2;
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
            this.cbGoogleCalendars.Location = new System.Drawing.Point(96, 42);
            this.cbGoogleCalendars.Name = "cbGoogleCalendars";
            this.cbGoogleCalendars.Size = new System.Drawing.Size(364, 21);
            this.cbGoogleCalendars.TabIndex = 1;
            this.cbGoogleCalendars.SelectedIndexChanged += new System.EventHandler(this.cbGoogleCalendars_SelectedIndexChanged);
            // 
            // gbOutlook
            // 
            this.gbOutlook.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbOutlook.Controls.Add(this.lOutlookCalendar);
            this.gbOutlook.Controls.Add(this.cbOutlookCalendars);
            this.gbOutlook.Controls.Add(this.rbOutlookEWS);
            this.gbOutlook.Controls.Add(this.rbOutlookDefaultMB);
            this.gbOutlook.Controls.Add(this.rbOutlookAltMB);
            this.gbOutlook.Controls.Add(this.gbEWS);
            this.gbOutlook.Controls.Add(this.ddMailboxName);
            this.gbOutlook.Location = new System.Drawing.Point(6, 6);
            this.gbOutlook.Name = "gbOutlook";
            this.gbOutlook.Size = new System.Drawing.Size(475, 149);
            this.gbOutlook.TabIndex = 15;
            this.gbOutlook.TabStop = false;
            this.gbOutlook.Text = "Outlook";
            // 
            // lOutlookCalendar
            // 
            this.lOutlookCalendar.AutoSize = true;
            this.lOutlookCalendar.Location = new System.Drawing.Point(9, 122);
            this.lOutlookCalendar.Name = "lOutlookCalendar";
            this.lOutlookCalendar.Size = new System.Drawing.Size(81, 13);
            this.lOutlookCalendar.TabIndex = 22;
            this.lOutlookCalendar.Text = "Select calendar";
            // 
            // cbOutlookCalendars
            // 
            this.cbOutlookCalendars.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cbOutlookCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutlookCalendars.FormattingEnabled = true;
            this.cbOutlookCalendars.Location = new System.Drawing.Point(96, 119);
            this.cbOutlookCalendars.Name = "cbOutlookCalendars";
            this.cbOutlookCalendars.Size = new System.Drawing.Size(364, 21);
            this.cbOutlookCalendars.TabIndex = 21;
            this.cbOutlookCalendars.SelectedIndexChanged += new System.EventHandler(this.cbOutlookCalendar_SelectedIndexChanged);
            // 
            // rbOutlookEWS
            // 
            this.rbOutlookEWS.AutoSize = true;
            this.rbOutlookEWS.Enabled = false;
            this.rbOutlookEWS.Location = new System.Drawing.Point(12, 92);
            this.rbOutlookEWS.Name = "rbOutlookEWS";
            this.rbOutlookEWS.Size = new System.Drawing.Size(143, 17);
            this.rbOutlookEWS.TabIndex = 19;
            this.rbOutlookEWS.Text = "Exchange Web Services";
            this.rbOutlookEWS.UseVisualStyleBackColor = true;
            this.rbOutlookEWS.CheckedChanged += new System.EventHandler(this.rbOutlookEWS_CheckedChanged);
            // 
            // rbOutlookDefaultMB
            // 
            this.rbOutlookDefaultMB.AutoSize = true;
            this.rbOutlookDefaultMB.Checked = true;
            this.rbOutlookDefaultMB.Location = new System.Drawing.Point(12, 19);
            this.rbOutlookDefaultMB.Name = "rbOutlookDefaultMB";
            this.rbOutlookDefaultMB.Size = new System.Drawing.Size(98, 17);
            this.rbOutlookDefaultMB.TabIndex = 18;
            this.rbOutlookDefaultMB.TabStop = true;
            this.rbOutlookDefaultMB.Text = "Default Mailbox";
            this.rbOutlookDefaultMB.UseVisualStyleBackColor = true;
            this.rbOutlookDefaultMB.CheckedChanged += new System.EventHandler(this.rbOutlookDefaultMB_CheckedChanged);
            // 
            // rbOutlookAltMB
            // 
            this.rbOutlookAltMB.AutoSize = true;
            this.rbOutlookAltMB.Location = new System.Drawing.Point(12, 42);
            this.rbOutlookAltMB.Name = "rbOutlookAltMB";
            this.rbOutlookAltMB.Size = new System.Drawing.Size(114, 17);
            this.rbOutlookAltMB.TabIndex = 17;
            this.rbOutlookAltMB.Text = "Alternative Mailbox";
            this.rbOutlookAltMB.UseVisualStyleBackColor = true;
            this.rbOutlookAltMB.CheckedChanged += new System.EventHandler(this.rbOutlookAltMB_CheckedChanged);
            // 
            // gbEWS
            // 
            this.gbEWS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbEWS.Controls.Add(this.txtEWSServerURL);
            this.gbEWS.Controls.Add(this.label8);
            this.gbEWS.Controls.Add(this.txtEWSPass);
            this.gbEWS.Controls.Add(this.label7);
            this.gbEWS.Controls.Add(this.txtEWSUser);
            this.gbEWS.Controls.Add(this.label6);
            this.gbEWS.Location = new System.Drawing.Point(221, 13);
            this.gbEWS.Name = "gbEWS";
            this.gbEWS.Size = new System.Drawing.Size(248, 96);
            this.gbEWS.TabIndex = 13;
            this.gbEWS.TabStop = false;
            this.gbEWS.Text = "EWS Settings";
            // 
            // txtEWSServerURL
            // 
            this.txtEWSServerURL.Location = new System.Drawing.Point(63, 67);
            this.txtEWSServerURL.Name = "txtEWSServerURL";
            this.txtEWSServerURL.Size = new System.Drawing.Size(176, 20);
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
            this.txtEWSPass.Location = new System.Drawing.Point(63, 41);
            this.txtEWSPass.Name = "txtEWSPass";
            this.txtEWSPass.Size = new System.Drawing.Size(176, 20);
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
            this.txtEWSUser.Location = new System.Drawing.Point(63, 15);
            this.txtEWSUser.Name = "txtEWSUser";
            this.txtEWSUser.Size = new System.Drawing.Size(176, 20);
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
            // ddMailboxName
            // 
            this.ddMailboxName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ddMailboxName.FormattingEnabled = true;
            this.ddMailboxName.Location = new System.Drawing.Point(31, 65);
            this.ddMailboxName.Name = "ddMailboxName";
            this.ddMailboxName.Size = new System.Drawing.Size(184, 21);
            this.ddMailboxName.TabIndex = 16;
            this.ddMailboxName.SelectedIndexChanged += new System.EventHandler(this.ddMailboxName_SelectedIndexChanged);
            // 
            // gbAppBehaviour
            // 
            this.gbAppBehaviour.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbAppBehaviour.Controls.Add(this.cbStartOnStartup);
            this.gbAppBehaviour.Controls.Add(this.cbShowBubbleTooltips);
            this.gbAppBehaviour.Controls.Add(this.cbMinimizeToTray);
            this.gbAppBehaviour.Controls.Add(this.cbStartInTray);
            this.gbAppBehaviour.Controls.Add(this.cbCreateFiles);
            this.gbAppBehaviour.Location = new System.Drawing.Point(218, 238);
            this.gbAppBehaviour.Name = "gbAppBehaviour";
            this.gbAppBehaviour.Size = new System.Drawing.Size(263, 122);
            this.gbAppBehaviour.TabIndex = 11;
            this.gbAppBehaviour.TabStop = false;
            this.gbAppBehaviour.Text = "Application Behaviour";
            // 
            // cbStartOnStartup
            // 
            this.cbStartOnStartup.AutoSize = true;
            this.cbStartOnStartup.Location = new System.Drawing.Point(9, 23);
            this.cbStartOnStartup.Name = "cbStartOnStartup";
            this.cbStartOnStartup.Size = new System.Drawing.Size(88, 17);
            this.cbStartOnStartup.TabIndex = 8;
            this.cbStartOnStartup.Text = "Start on login";
            this.cbStartOnStartup.UseVisualStyleBackColor = true;
            this.cbStartOnStartup.CheckedChanged += new System.EventHandler(this.cbStartOnStartup_CheckedChanged);
            // 
            // cbShowBubbleTooltips
            // 
            this.cbShowBubbleTooltips.Location = new System.Drawing.Point(9, 95);
            this.cbShowBubbleTooltips.Name = "cbShowBubbleTooltips";
            this.cbShowBubbleTooltips.Size = new System.Drawing.Size(259, 24);
            this.cbShowBubbleTooltips.TabIndex = 7;
            this.cbShowBubbleTooltips.Text = "Show Bubble Tooltip in Taskbar when Syncing";
            this.cbShowBubbleTooltips.UseVisualStyleBackColor = true;
            this.cbShowBubbleTooltips.CheckedChanged += new System.EventHandler(this.cbShowBubbleTooltipsCheckedChanged);
            // 
            // cbMinimizeToTray
            // 
            this.cbMinimizeToTray.Location = new System.Drawing.Point(9, 57);
            this.cbMinimizeToTray.Name = "cbMinimizeToTray";
            this.cbMinimizeToTray.Size = new System.Drawing.Size(104, 24);
            this.cbMinimizeToTray.TabIndex = 0;
            this.cbMinimizeToTray.Text = "Minimize to Tray";
            this.cbMinimizeToTray.UseVisualStyleBackColor = true;
            this.cbMinimizeToTray.CheckedChanged += new System.EventHandler(this.cbMinimizeToTrayCheckedChanged);
            // 
            // cbStartInTray
            // 
            this.cbStartInTray.Location = new System.Drawing.Point(9, 38);
            this.cbStartInTray.Name = "cbStartInTray";
            this.cbStartInTray.Size = new System.Drawing.Size(104, 24);
            this.cbStartInTray.TabIndex = 1;
            this.cbStartInTray.Text = "Start in Tray";
            this.cbStartInTray.UseVisualStyleBackColor = true;
            this.cbStartInTray.CheckedChanged += new System.EventHandler(this.cbStartInTrayCheckedChanged);
            // 
            // cbCreateFiles
            // 
            this.cbCreateFiles.Location = new System.Drawing.Point(9, 76);
            this.cbCreateFiles.Name = "cbCreateFiles";
            this.cbCreateFiles.Size = new System.Drawing.Size(235, 24);
            this.cbCreateFiles.TabIndex = 7;
            this.cbCreateFiles.Text = "Create CSV files of calendar entries";
            this.cbCreateFiles.UseVisualStyleBackColor = true;
            this.cbCreateFiles.CheckedChanged += new System.EventHandler(this.cbCreateFiles_CheckedChanged);
            // 
            // bSave
            // 
            this.bSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSave.Location = new System.Drawing.Point(391, 494);
            this.bSave.Name = "bSave";
            this.bSave.Size = new System.Drawing.Size(75, 31);
            this.bSave.TabIndex = 8;
            this.bSave.Text = "Save";
            this.bSave.UseVisualStyleBackColor = true;
            this.bSave.Click += new System.EventHandler(this.Save_Click);
            // 
            // gbSyncOptions
            // 
            this.gbSyncOptions.Controls.Add(this.cbOutlookPush);
            this.gbSyncOptions.Controls.Add(this.cbMergeItems);
            this.gbSyncOptions.Controls.Add(this.syncDirection);
            this.gbSyncOptions.Controls.Add(this.cbDisableDeletion);
            this.gbSyncOptions.Controls.Add(this.lMiscOptions);
            this.gbSyncOptions.Controls.Add(this.cbConfirmOnDelete);
            this.gbSyncOptions.Controls.Add(this.cbAddReminders);
            this.gbSyncOptions.Controls.Add(this.lAttributes);
            this.gbSyncOptions.Controls.Add(this.cbAddAttendees);
            this.gbSyncOptions.Controls.Add(this.cbIntervalUnit);
            this.gbSyncOptions.Controls.Add(this.cbAddDescription);
            this.gbSyncOptions.Controls.Add(this.tbInterval);
            this.gbSyncOptions.Controls.Add(this.label1);
            this.gbSyncOptions.Controls.Add(this.tbDaysInTheFuture);
            this.gbSyncOptions.Controls.Add(this.tbDaysInThePast);
            this.gbSyncOptions.Controls.Add(this.lDaysInFuture);
            this.gbSyncOptions.Controls.Add(this.lDaysInPast);
            this.gbSyncOptions.Controls.Add(this.lDateRange);
            this.gbSyncOptions.Location = new System.Drawing.Point(6, 238);
            this.gbSyncOptions.Name = "gbSyncOptions";
            this.gbSyncOptions.Size = new System.Drawing.Size(206, 298);
            this.gbSyncOptions.TabIndex = 0;
            this.gbSyncOptions.TabStop = false;
            this.gbSyncOptions.Text = "Sync Options";
            // 
            // cbOutlookPush
            // 
            this.cbOutlookPush.AutoSize = true;
            this.cbOutlookPush.Enabled = false;
            this.cbOutlookPush.Location = new System.Drawing.Point(15, 119);
            this.cbOutlookPush.Name = "cbOutlookPush";
            this.cbOutlookPush.Size = new System.Drawing.Size(191, 17);
            this.cbOutlookPush.TabIndex = 15;
            this.cbOutlookPush.Text = "Push Outlook changes immediately";
            this.cbOutlookPush.UseVisualStyleBackColor = true;
            this.cbOutlookPush.CheckedChanged += new System.EventHandler(this.cbOutlookPush_CheckedChanged);
            // 
            // cbMergeItems
            // 
            this.cbMergeItems.Location = new System.Drawing.Point(14, 234);
            this.cbMergeItems.Name = "cbMergeItems";
            this.cbMergeItems.Size = new System.Drawing.Size(152, 17);
            this.cbMergeItems.TabIndex = 14;
            this.cbMergeItems.Text = "Merge with existing entries";
            this.cbMergeItems.UseVisualStyleBackColor = true;
            this.cbMergeItems.CheckedChanged += new System.EventHandler(this.cbMergeItems_CheckedChanged);
            // 
            // syncDirection
            // 
            this.syncDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.syncDirection.FormattingEnabled = true;
            this.syncDirection.Location = new System.Drawing.Point(12, 19);
            this.syncDirection.Name = "syncDirection";
            this.syncDirection.Size = new System.Drawing.Size(186, 21);
            this.syncDirection.TabIndex = 13;
            this.syncDirection.SelectedIndexChanged += new System.EventHandler(this.syncDirection_SelectedIndexChanged);
            // 
            // cbDisableDeletion
            // 
            this.cbDisableDeletion.Location = new System.Drawing.Point(14, 253);
            this.cbDisableDeletion.Name = "cbDisableDeletion";
            this.cbDisableDeletion.Size = new System.Drawing.Size(106, 17);
            this.cbDisableDeletion.TabIndex = 8;
            this.cbDisableDeletion.Text = "Disable deletions";
            this.cbDisableDeletion.UseVisualStyleBackColor = true;
            this.cbDisableDeletion.CheckedChanged += new System.EventHandler(this.cbDisableDeletion_CheckedChanged);
            // 
            // lMiscOptions
            // 
            this.lMiscOptions.Location = new System.Drawing.Point(3, 217);
            this.lMiscOptions.Name = "lMiscOptions";
            this.lMiscOptions.Size = new System.Drawing.Size(103, 14);
            this.lMiscOptions.TabIndex = 12;
            this.lMiscOptions.Text = "Miscellenous:-";
            // 
            // cbConfirmOnDelete
            // 
            this.cbConfirmOnDelete.Location = new System.Drawing.Point(14, 271);
            this.cbConfirmOnDelete.Name = "cbConfirmOnDelete";
            this.cbConfirmOnDelete.Size = new System.Drawing.Size(111, 17);
            this.cbConfirmOnDelete.TabIndex = 9;
            this.cbConfirmOnDelete.Text = "Confirm deletions";
            this.cbConfirmOnDelete.UseVisualStyleBackColor = true;
            this.cbConfirmOnDelete.CheckedChanged += new System.EventHandler(this.cbConfirmOnDelete_CheckedChanged);
            // 
            // cbAddReminders
            // 
            this.cbAddReminders.Location = new System.Drawing.Point(15, 194);
            this.cbAddReminders.Name = "cbAddReminders";
            this.cbAddReminders.Size = new System.Drawing.Size(80, 17);
            this.cbAddReminders.TabIndex = 8;
            this.cbAddReminders.Text = "Reminders";
            this.cbAddReminders.UseVisualStyleBackColor = true;
            this.cbAddReminders.CheckedChanged += new System.EventHandler(this.CbAddRemindersCheckedChanged);
            // 
            // lAttributes
            // 
            this.lAttributes.Location = new System.Drawing.Point(3, 142);
            this.lAttributes.Name = "lAttributes";
            this.lAttributes.Size = new System.Drawing.Size(103, 14);
            this.lAttributes.TabIndex = 11;
            this.lAttributes.Text = "Attributes included:-";
            // 
            // cbAddAttendees
            // 
            this.cbAddAttendees.Location = new System.Drawing.Point(15, 176);
            this.cbAddAttendees.Name = "cbAddAttendees";
            this.cbAddAttendees.Size = new System.Drawing.Size(80, 17);
            this.cbAddAttendees.TabIndex = 6;
            this.cbAddAttendees.Text = "Attendees";
            this.cbAddAttendees.UseVisualStyleBackColor = true;
            this.cbAddAttendees.CheckedChanged += new System.EventHandler(this.cbAddAttendees_CheckedChanged);
            // 
            // cbIntervalUnit
            // 
            this.cbIntervalUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbIntervalUnit.FormattingEnabled = true;
            this.cbIntervalUnit.Items.AddRange(new object[] {
            "Minutes",
            "Hours"});
            this.cbIntervalUnit.Location = new System.Drawing.Point(114, 92);
            this.cbIntervalUnit.Name = "cbIntervalUnit";
            this.cbIntervalUnit.Size = new System.Drawing.Size(84, 21);
            this.cbIntervalUnit.TabIndex = 10;
            this.cbIntervalUnit.SelectedIndexChanged += new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
            // 
            // cbAddDescription
            // 
            this.cbAddDescription.Location = new System.Drawing.Point(15, 158);
            this.cbAddDescription.Name = "cbAddDescription";
            this.cbAddDescription.Size = new System.Drawing.Size(80, 17);
            this.cbAddDescription.TabIndex = 7;
            this.cbAddDescription.Text = "Description";
            this.cbAddDescription.UseVisualStyleBackColor = true;
            this.cbAddDescription.CheckedChanged += new System.EventHandler(this.CbAddDescriptionCheckedChanged);
            // 
            // tbInterval
            // 
            this.tbInterval.Location = new System.Drawing.Point(70, 93);
            this.tbInterval.Maximum = new decimal(new int[] {
            120,
            0,
            0,
            0});
            this.tbInterval.Name = "tbInterval";
            this.tbInterval.Size = new System.Drawing.Size(40, 20);
            this.tbInterval.TabIndex = 9;
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
            this.label1.Location = new System.Drawing.Point(6, 95);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 14);
            this.label1.TabIndex = 8;
            this.label1.Text = "Interval:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // tbDaysInTheFuture
            // 
            this.tbDaysInTheFuture.Location = new System.Drawing.Point(70, 69);
            this.tbDaysInTheFuture.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.tbDaysInTheFuture.Name = "tbDaysInTheFuture";
            this.tbDaysInTheFuture.Size = new System.Drawing.Size(40, 20);
            this.tbDaysInTheFuture.TabIndex = 7;
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
            this.tbDaysInThePast.Location = new System.Drawing.Point(70, 45);
            this.tbDaysInThePast.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.tbDaysInThePast.Name = "tbDaysInThePast";
            this.tbDaysInThePast.Size = new System.Drawing.Size(40, 20);
            this.tbDaysInThePast.TabIndex = 5;
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
            this.lDaysInFuture.Location = new System.Drawing.Point(111, 71);
            this.lDaysInFuture.Name = "lDaysInFuture";
            this.lDaysInFuture.Size = new System.Drawing.Size(104, 20);
            this.lDaysInFuture.TabIndex = 0;
            this.lDaysInFuture.Text = "days in the future";
            // 
            // lDaysInPast
            // 
            this.lDaysInPast.Location = new System.Drawing.Point(111, 48);
            this.lDaysInPast.Name = "lDaysInPast";
            this.lDaysInPast.Size = new System.Drawing.Size(87, 18);
            this.lDaysInPast.TabIndex = 0;
            this.lDaysInPast.Text = "days in the past";
            // 
            // lDateRange
            // 
            this.lDateRange.Location = new System.Drawing.Point(6, 48);
            this.lDateRange.Name = "lDateRange";
            this.lDateRange.Size = new System.Drawing.Size(66, 14);
            this.lDateRange.TabIndex = 6;
            this.lDateRange.Text = "Date range:";
            this.lDateRange.TextAlign = System.Drawing.ContentAlignment.TopRight;
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
            this.notifyIcon1.Text = "OutlookGoogleSync";
            this.notifyIcon1.Click += new System.EventHandler(this.NotifyIcon1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 592);
            this.Controls.Add(this.tabSettings);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(535, 630);
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Outlook Google Calendar Sync";
            this.Resize += new System.EventHandler(this.MainFormResize);
            this.tabSettings.ResumeLayout(false);
            this.tabPage_Sync.ResumeLayout(false);
            this.tabPage_Sync.PerformLayout();
            this.tabPage_Settings.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.gbGoogle.ResumeLayout(false);
            this.gbGoogle.PerformLayout();
            this.gbOutlook.ResumeLayout(false);
            this.gbOutlook.PerformLayout();
            this.gbEWS.ResumeLayout(false);
            this.gbEWS.PerformLayout();
            this.gbAppBehaviour.ResumeLayout(false);
            this.gbAppBehaviour.PerformLayout();
            this.gbSyncOptions.ResumeLayout(false);
            this.gbSyncOptions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInTheFuture)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInThePast)).EndInit();
            this.tabPage_About.ResumeLayout(false);
            this.tabPage_About.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).EndInit();
            this.ResumeLayout(false);

        }
        private System.Windows.Forms.CheckBox cbAddReminders;
        private System.Windows.Forms.CheckBox cbAddDescription;
        private System.Windows.Forms.CheckBox cbShowBubbleTooltips;
        private System.Windows.Forms.CheckBox cbMinimizeToTray;
        private System.Windows.Forms.CheckBox cbStartInTray;
        private System.Windows.Forms.CheckBox cbDisableDeletion;
        private System.Windows.Forms.CheckBox cbConfirmOnDelete;
        private System.Windows.Forms.GroupBox gbAppBehaviour;
        private System.Windows.Forms.LinkLabel lAboutURL;
        private System.Windows.Forms.TabPage tabPage_About;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.Label lAboutMain;
        private System.Windows.Forms.CheckBox cbAddAttendees;
        private System.Windows.Forms.TextBox LogBox;
        private System.Windows.Forms.GroupBox gbGoogle;
        private System.Windows.Forms.Label lGoogleCalendar;
        private System.Windows.Forms.Label lDaysInPast;
        private System.Windows.Forms.Label lDaysInFuture;
        private System.Windows.Forms.GroupBox gbSyncOptions;
        private System.Windows.Forms.Button bSave;
        private System.Windows.Forms.Button bSyncNow;
        private System.Windows.Forms.TabPage tabPage_Sync;
        private System.Windows.Forms.Button bGetGoogleCalendars;
        private System.Windows.Forms.GroupBox gbOutlook;
        private System.Windows.Forms.GroupBox gbEWS;
        private System.Windows.Forms.TextBox txtEWSPass;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtEWSUser;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lLastSync;
        private System.Windows.Forms.Label lNextSync;
        private System.Windows.Forms.Label lNextSyncVal;
        private System.Windows.Forms.Label lLastSyncVal;
        private System.Windows.Forms.ComboBox ddMailboxName;
        private System.Windows.Forms.TextBox txtEWSServerURL;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.RadioButton rbOutlookDefaultMB;
        private System.Windows.Forms.RadioButton rbOutlookAltMB;
        private System.Windows.Forms.RadioButton rbOutlookEWS;
        private System.Windows.Forms.ComboBox cbOutlookCalendars;
        private System.Windows.Forms.Label lOutlookCalendar;
        private System.Windows.Forms.Label lGoogleHelp;
        private System.Windows.Forms.NumericUpDown tbDaysInThePast;
        private System.Windows.Forms.Label lDateRange;
        private System.Windows.Forms.NumericUpDown tbDaysInTheFuture;
        private System.Windows.Forms.NumericUpDown tbInterval;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbIntervalUnit;
        private System.Windows.Forms.Label lAttributes;
        private System.Windows.Forms.Label lMiscOptions;
        private System.Windows.Forms.ComboBox cbGoogleCalendars;
        private System.Windows.Forms.CheckBox cbVerboseOutput;
        private System.Windows.Forms.CheckBox cbCreateFiles;
        private System.Windows.Forms.ComboBox syncDirection;
        private System.Windows.Forms.TextBox tbHelp;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox cbMergeItems;
        public System.Windows.Forms.TabControl tabSettings;
        public System.Windows.Forms.TabPage tabPage_Settings;
        private System.Windows.Forms.PictureBox pbDonate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cbStartOnStartup;
        private System.Windows.Forms.CheckBox cbOutlookPush;
	}
}
