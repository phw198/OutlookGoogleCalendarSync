﻿namespace OutlookGoogleCalendarSync.Forms {
    partial class Splash {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Splash));
            this.panel1 = new System.Windows.Forms.Panel();
            this.pHideTwitterBt = new System.Windows.Forms.Panel();
            this.cbHideSplash = new System.Windows.Forms.CheckBox();
            this.lSyncCount = new System.Windows.Forms.Label();
            this.pbSocialTwitterFollow = new System.Windows.Forms.PictureBox();
            this.lVersion = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pHideDonateBt = new System.Windows.Forms.Panel();
            this.pbDonate = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbSocialTwitterFollow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.pHideTwitterBt);
            this.panel1.Controls.Add(this.cbHideSplash);
            this.panel1.Controls.Add(this.lSyncCount);
            this.panel1.Controls.Add(this.pbSocialTwitterFollow);
            this.panel1.Controls.Add(this.lVersion);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.pHideDonateBt);
            this.panel1.Controls.Add(this.pbDonate);
            this.panel1.Location = new System.Drawing.Point(4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(463, 303);
            this.panel1.TabIndex = 9;
            // 
            // pHideTwitterBt
            // 
            this.pHideTwitterBt.Location = new System.Drawing.Point(162, 277);
            this.pHideTwitterBt.Name = "pHideTwitterBt";
            this.pHideTwitterBt.Size = new System.Drawing.Size(147, 26);
            this.pHideTwitterBt.TabIndex = 58;
            // 
            // cbHideSplash
            // 
            this.cbHideSplash.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cbHideSplash.AutoSize = true;
            this.cbHideSplash.Location = new System.Drawing.Point(8, 283);
            this.cbHideSplash.Name = "cbHideSplash";
            this.cbHideSplash.Size = new System.Drawing.Size(120, 17);
            this.cbHideSplash.TabIndex = 56;
            this.cbHideSplash.Text = "Hide Splash Screen";
            this.cbHideSplash.UseVisualStyleBackColor = true;
            this.cbHideSplash.CheckedChanged += new System.EventHandler(this.cbHideSplash_CheckedChanged);
            // 
            // lSyncCount
            // 
            this.lSyncCount.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lSyncCount.AutoSize = true;
            this.lSyncCount.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lSyncCount.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lSyncCount.Location = new System.Drawing.Point(149, 198);
            this.lSyncCount.Name = "lSyncCount";
            this.lSyncCount.Size = new System.Drawing.Size(173, 15);
            this.lSyncCount.TabIndex = 55;
            this.lSyncCount.Text = "{syncs} Syncs && Counting...";
            this.lSyncCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pbSocialTwitterFollow
            // 
            this.pbSocialTwitterFollow.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.pbSocialTwitterFollow.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbSocialTwitterFollow.Image = global::OutlookGoogleCalendarSync.Properties.Resources.bluesky_follow;
            this.pbSocialTwitterFollow.Location = new System.Drawing.Point(168, 251);
            this.pbSocialTwitterFollow.Name = "pbSocialTwitterFollow";
            this.pbSocialTwitterFollow.Size = new System.Drawing.Size(135, 27);
            this.pbSocialTwitterFollow.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbSocialTwitterFollow.TabIndex = 54;
            this.pbSocialTwitterFollow.TabStop = false;
            this.pbSocialTwitterFollow.Click += new System.EventHandler(this.pbSocialTwitterFollow_Click);
            // 
            // lVersion
            // 
            this.lVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lVersion.Location = new System.Drawing.Point(339, 284);
            this.lVersion.Name = "lVersion";
            this.lVersion.Size = new System.Drawing.Size(116, 13);
            this.lVersion.TabIndex = 12;
            this.lVersion.Text = "Version";
            this.lVersion.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Black", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.DimGray;
            this.label1.Location = new System.Drawing.Point(126, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(218, 33);
            this.label1.TabIndex = 10;
            this.label1.Text = "Outlook Google";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::OutlookGoogleCalendarSync.Properties.Resources.ogcs128x128_animated;
            this.pictureBox1.Location = new System.Drawing.Point(170, 39);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(130, 130);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Black", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DimGray;
            this.label2.Location = new System.Drawing.Point(133, 164);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(205, 33);
            this.label2.TabIndex = 11;
            this.label2.Text = "Calendar Sync";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // pHideDonateBt
            // 
            this.pHideDonateBt.Location = new System.Drawing.Point(180, 242);
            this.pHideDonateBt.Name = "pHideDonateBt";
            this.pHideDonateBt.Size = new System.Drawing.Size(110, 35);
            this.pHideDonateBt.TabIndex = 57;
            // 
            // pbDonate
            // 
            this.pbDonate.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pbDonate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbDonate.Image = global::OutlookGoogleCalendarSync.Properties.Resources.paypalDonate;
            this.pbDonate.Location = new System.Drawing.Point(198, 221);
            this.pbDonate.Name = "pbDonate";
            this.pbDonate.Size = new System.Drawing.Size(74, 21);
            this.pbDonate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pbDonate.TabIndex = 8;
            this.pbDonate.TabStop = false;
            this.pbDonate.Click += new System.EventHandler(this.pbDonate_Click);
            // 
            // Splash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Peru;
            this.ClientSize = new System.Drawing.Size(471, 310);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Splash";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Outlook Google Calendar Sync";
            this.TopMost = true;
            this.Shown += new System.EventHandler(this.Splash_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbSocialTwitterFollow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pbDonate;
        private System.Windows.Forms.Label lVersion;
        private System.Windows.Forms.PictureBox pbSocialTwitterFollow;
        private System.Windows.Forms.Label lSyncCount;
        private System.Windows.Forms.CheckBox cbHideSplash;
        private System.Windows.Forms.Panel pHideDonateBt;
        private System.Windows.Forms.Panel pHideTwitterBt;
    }
}