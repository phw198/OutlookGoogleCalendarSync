namespace OutlookGoogleCalendarSync.Forms {
    partial class UpdateInfo {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (this.optionChosen == System.Windows.Forms.DialogResult.Yes && !this.AwaitingRestart) {
                this.Visible = true;
                disposing = false;
            }
            try {
                if (disposing && (components != null)) {
                    components.Dispose();
                }
                base.Dispose(disposing);
            } catch {
                //WebBrowser COM object is already disposed of, but we don't care at this point.
            }
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UpdateInfo));
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.btUpgrade = new System.Windows.Forms.Button();
            this.wbPanel = new System.Windows.Forms.Panel();
            this.llViewOnGithub = new System.Windows.Forms.LinkLabel();
            this.btLater = new System.Windows.Forms.Button();
            this.lSummary = new System.Windows.Forms.Label();
            this.lTitle = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btSkipVersion = new System.Windows.Forms.Button();
            this.wbPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // webBrowser
            // 
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.Location = new System.Drawing.Point(0, 0);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new System.Drawing.Size(575, 238);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.WebBrowserShortcutsEnabled = false;
            this.webBrowser.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.webBrowser_Navigating);
            // 
            // btUpgrade
            // 
            this.btUpgrade.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btUpgrade.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.btUpgrade.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btUpgrade.Location = new System.Drawing.Point(497, 329);
            this.btUpgrade.Name = "btUpgrade";
            this.btUpgrade.Size = new System.Drawing.Size(94, 30);
            this.btUpgrade.TabIndex = 1;
            this.btUpgrade.Text = "Upgrade";
            this.btUpgrade.UseVisualStyleBackColor = true;
            this.btUpgrade.Click += new System.EventHandler(this.btUpgrade_Click);
            // 
            // wbPanel
            // 
            this.wbPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wbPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.wbPanel.Controls.Add(this.llViewOnGithub);
            this.wbPanel.Controls.Add(this.webBrowser);
            this.wbPanel.Location = new System.Drawing.Point(15, 80);
            this.wbPanel.Name = "wbPanel";
            this.wbPanel.Size = new System.Drawing.Size(577, 240);
            this.wbPanel.TabIndex = 2;
            // 
            // llViewOnGithub
            // 
            this.llViewOnGithub.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.llViewOnGithub.AutoSize = true;
            this.llViewOnGithub.BackColor = System.Drawing.Color.White;
            this.llViewOnGithub.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.llViewOnGithub.Location = new System.Drawing.Point(176, 38);
            this.llViewOnGithub.Name = "llViewOnGithub";
            this.llViewOnGithub.Size = new System.Drawing.Size(201, 20);
            this.llViewOnGithub.TabIndex = 9;
            this.llViewOnGithub.TabStop = true;
            this.llViewOnGithub.Text = "View Release Notes Online";
            this.llViewOnGithub.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.llViewOnGithub.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llViewOnGithub_LinkClicked);
            // 
            // btLater
            // 
            this.btLater.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btLater.DialogResult = System.Windows.Forms.DialogResult.No;
            this.btLater.Location = new System.Drawing.Point(397, 329);
            this.btLater.Name = "btLater";
            this.btLater.Size = new System.Drawing.Size(94, 30);
            this.btLater.TabIndex = 3;
            this.btLater.Text = "Later";
            this.btLater.UseVisualStyleBackColor = true;
            // 
            // lSummary
            // 
            this.lSummary.AutoSize = true;
            this.lSummary.Location = new System.Drawing.Point(12, 35);
            this.lSummary.Name = "lSummary";
            this.lSummary.Size = new System.Drawing.Size(214, 13);
            this.lSummary.TabIndex = 4;
            this.lSummary.Text = "Would you like to upgrade to v1.2.3.4 now?";
            // 
            // lTitle
            // 
            this.lTitle.AutoSize = true;
            this.lTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lTitle.Location = new System.Drawing.Point(12, 9);
            this.lTitle.Name = "lTitle";
            this.lTitle.Size = new System.Drawing.Size(275, 18);
            this.lTitle.TabIndex = 5;
            this.lTitle.Text = "A new release of OGCS is available";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Release Notes:";
            // 
            // btSkipVersion
            // 
            this.btSkipVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btSkipVersion.DialogResult = System.Windows.Forms.DialogResult.Ignore;
            this.btSkipVersion.Location = new System.Drawing.Point(16, 329);
            this.btSkipVersion.Name = "btSkipVersion";
            this.btSkipVersion.Size = new System.Drawing.Size(104, 30);
            this.btSkipVersion.TabIndex = 7;
            this.btSkipVersion.Text = "Skip This Version";
            this.btSkipVersion.UseVisualStyleBackColor = true;
            this.btSkipVersion.Click += new System.EventHandler(this.btSkipVersion_Click);
            // 
            // UpdateInfo
            // 
            this.AcceptButton = this.btUpgrade;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btLater;
            this.ClientSize = new System.Drawing.Size(604, 367);
            this.Controls.Add(this.btSkipVersion);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lTitle);
            this.Controls.Add(this.lSummary);
            this.Controls.Add(this.btLater);
            this.Controls.Add(this.wbPanel);
            this.Controls.Add(this.btUpgrade);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(510, 333);
            this.Name = "UpdateInfo";
            this.Text = "OGCS Update Available";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.UpdateInfo_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.UpdateInfo_FormClosed);
            this.wbPanel.ResumeLayout(false);
            this.wbPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Button btUpgrade;
        private System.Windows.Forms.Panel wbPanel;
        private System.Windows.Forms.Button btLater;
        private System.Windows.Forms.Label lSummary;
        private System.Windows.Forms.Label lTitle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btSkipVersion;
        private System.Windows.Forms.LinkLabel llViewOnGithub;
    }
}