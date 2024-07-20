
namespace OutlookGoogleCalendarSync.Forms {
    partial class ProfileManage {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProfileManage));
            this.txtProfileName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btOK = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.pbDonate = new System.Windows.Forms.PictureBox();
            this.panelDonationNote = new System.Windows.Forms.Panel();
            this.tbDonate = new System.Windows.Forms.RichTextBox();
            this.tbProfileInfo = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).BeginInit();
            this.panelDonationNote.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtProfileName
            // 
            this.txtProfileName.Location = new System.Drawing.Point(166, 22);
            this.txtProfileName.Name = "txtProfileName";
            this.txtProfileName.Size = new System.Drawing.Size(231, 20);
            this.txtProfileName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(90, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Profile Name:";
            // 
            // btOK
            // 
            this.btOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btOK.Location = new System.Drawing.Point(403, 237);
            this.btOK.Name = "btOK";
            this.btOK.Size = new System.Drawing.Size(91, 23);
            this.btOK.TabIndex = 2;
            this.btOK.Text = "OK";
            this.btOK.UseVisualStyleBackColor = true;
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btCancel.Location = new System.Drawing.Point(308, 237);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(89, 23);
            this.btCancel.TabIndex = 3;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            // 
            // pbDonate
            // 
            this.pbDonate.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pbDonate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbDonate.Image = global::OutlookGoogleCalendarSync.Properties.Resources.paypalDonate;
            this.pbDonate.Location = new System.Drawing.Point(195, 76);
            this.pbDonate.Name = "pbDonate";
            this.pbDonate.Size = new System.Drawing.Size(76, 24);
            this.pbDonate.TabIndex = 9;
            this.pbDonate.TabStop = false;
            this.pbDonate.Click += new System.EventHandler(this.pbDonate_Click);
            // 
            // panelDonationNote
            // 
            this.panelDonationNote.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.panelDonationNote.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.panelDonationNote.Controls.Add(this.pbDonate);
            this.panelDonationNote.Controls.Add(this.tbDonate);
            this.panelDonationNote.Location = new System.Drawing.Point(29, 112);
            this.panelDonationNote.Name = "panelDonationNote";
            this.panelDonationNote.Padding = new System.Windows.Forms.Padding(10);
            this.panelDonationNote.Size = new System.Drawing.Size(451, 113);
            this.panelDonationNote.TabIndex = 36;
            // 
            // tbDonate
            // 
            this.tbDonate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.tbDonate.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbDonate.Cursor = System.Windows.Forms.Cursors.Default;
            this.tbDonate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbDonate.Font = new System.Drawing.Font("Comic Sans MS", 9.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbDonate.Location = new System.Drawing.Point(10, 10);
            this.tbDonate.Name = "tbDonate";
            this.tbDonate.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.tbDonate.Size = new System.Drawing.Size(431, 93);
            this.tbDonate.TabIndex = 2;
            this.tbDonate.TabStop = false;
            this.tbDonate.Text = "Many applications would put this kind of feature behind a paywall or license.\n\nNo" +
    "t OCGCS. The aim has always been to keep it free...though you can show your appr" +
    "eciation via donation!";
            // 
            // tbProfileInfo
            // 
            this.tbProfileInfo.BackColor = System.Drawing.SystemColors.Control;
            this.tbProfileInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbProfileInfo.Location = new System.Drawing.Point(51, 55);
            this.tbProfileInfo.Multiline = true;
            this.tbProfileInfo.Name = "tbProfileInfo";
            this.tbProfileInfo.Size = new System.Drawing.Size(416, 42);
            this.tbProfileInfo.TabIndex = 37;
            this.tbProfileInfo.Text = "Each Profile will store the configuration defined under the \"Sync Settings\" tab.\r" +
    "\n\r\nThis allows for more than one calendar to be synced, each with their own sche" +
    "dule.";
            // 
            // ProfileManage
            // 
            this.AcceptButton = this.btOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btCancel;
            this.ClientSize = new System.Drawing.Size(506, 272);
            this.Controls.Add(this.tbProfileInfo);
            this.Controls.Add(this.panelDonationNote);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btOK);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtProfileName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProfileManage";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Profile Management";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProfileManage_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).EndInit();
            this.panelDonationNote.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtProfileName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btOK;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.PictureBox pbDonate;
        private System.Windows.Forms.Panel panelDonationNote;
        private System.Windows.Forms.RichTextBox tbDonate;
        private System.Windows.Forms.TextBox tbProfileInfo;
    }
}