
namespace OutlookGoogleCalendarSync.Forms {
    partial class MsOauthConsent {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MsOauthConsent));
            this.rbBlocked = new System.Windows.Forms.RadioButton();
            this.rbDoubts = new System.Windows.Forms.RadioButton();
            this.rbJustificationGiven = new System.Windows.Forms.RadioButton();
            this.rbAdminGrant = new System.Windows.Forms.RadioButton();
            this.rbEndOfRoad = new System.Windows.Forms.RadioButton();
            this.rbDunno = new System.Windows.Forms.RadioButton();
            this.btOK = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbBlocked
            // 
            this.rbBlocked.AutoSize = true;
            this.rbBlocked.Location = new System.Drawing.Point(31, 88);
            this.rbBlocked.Name = "rbBlocked";
            this.rbBlocked.Size = new System.Drawing.Size(205, 17);
            this.rbBlocked.TabIndex = 1;
            this.rbBlocked.TabStop = true;
            this.rbBlocked.Text = "My corporate IT have blocked access";
            this.rbBlocked.UseVisualStyleBackColor = true;
            this.rbBlocked.CheckedChanged += new System.EventHandler(this.rbBlocked_CheckedChanged);
            // 
            // rbDoubts
            // 
            this.rbDoubts.AutoSize = true;
            this.rbDoubts.Location = new System.Drawing.Point(31, 65);
            this.rbDoubts.Name = "rbDoubts";
            this.rbDoubts.Size = new System.Drawing.Size(269, 17);
            this.rbDoubts.TabIndex = 2;
            this.rbDoubts.TabStop = true;
            this.rbDoubts.Text = "I\'ve doubts about giving this application that access";
            this.rbDoubts.UseVisualStyleBackColor = true;
            this.rbDoubts.Click += new System.EventHandler(this.rbDoubts_Click);
            // 
            // rbJustificationGiven
            // 
            this.rbJustificationGiven.AutoSize = true;
            this.rbJustificationGiven.Enabled = false;
            this.rbJustificationGiven.Location = new System.Drawing.Point(3, 3);
            this.rbJustificationGiven.Name = "rbJustificationGiven";
            this.rbJustificationGiven.Size = new System.Drawing.Size(258, 17);
            this.rbJustificationGiven.TabIndex = 3;
            this.rbJustificationGiven.TabStop = true;
            this.rbJustificationGiven.Text = "I\'ve requested access and provided a justification";
            this.rbJustificationGiven.UseVisualStyleBackColor = true;
            this.rbJustificationGiven.Click += new System.EventHandler(this.rbJustificationGiven_Click);
            // 
            // rbAdminGrant
            // 
            this.rbAdminGrant.AutoSize = true;
            this.rbAdminGrant.Enabled = false;
            this.rbAdminGrant.Location = new System.Drawing.Point(3, 26);
            this.rbAdminGrant.Name = "rbAdminGrant";
            this.rbAdminGrant.Size = new System.Drawing.Size(249, 17);
            this.rbAdminGrant.TabIndex = 4;
            this.rbAdminGrant.TabStop = true;
            this.rbAdminGrant.Text = "I\'m going to ask an IT admin to grant permission";
            this.rbAdminGrant.UseVisualStyleBackColor = true;
            this.rbAdminGrant.Click += new System.EventHandler(this.rbAdminGrant_Click);
            // 
            // rbEndOfRoad
            // 
            this.rbEndOfRoad.AutoSize = true;
            this.rbEndOfRoad.Enabled = false;
            this.rbEndOfRoad.Location = new System.Drawing.Point(3, 49);
            this.rbEndOfRoad.Name = "rbEndOfRoad";
            this.rbEndOfRoad.Size = new System.Drawing.Size(236, 17);
            this.rbEndOfRoad.TabIndex = 5;
            this.rbEndOfRoad.TabStop = true;
            this.rbEndOfRoad.Text = "There\'s nothing I can do, it\'s out of my hands";
            this.rbEndOfRoad.UseVisualStyleBackColor = true;
            this.rbEndOfRoad.Click += new System.EventHandler(this.rbEndOfRoad_Click);
            // 
            // rbDunno
            // 
            this.rbDunno.AutoSize = true;
            this.rbDunno.Location = new System.Drawing.Point(31, 180);
            this.rbDunno.Name = "rbDunno";
            this.rbDunno.Size = new System.Drawing.Size(126, 17);
            this.rbDunno.TabIndex = 6;
            this.rbDunno.TabStop = true;
            this.rbDunno.Text = "Not sure - I\'ll try again";
            this.rbDunno.UseVisualStyleBackColor = true;
            this.rbDunno.Click += new System.EventHandler(this.rbDunno_Click);
            // 
            // btOK
            // 
            this.btOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btOK.Enabled = false;
            this.btOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btOK.Location = new System.Drawing.Point(325, 201);
            this.btOK.Name = "btOK";
            this.btOK.Size = new System.Drawing.Size(75, 23);
            this.btOK.TabIndex = 7;
            this.btOK.Text = "OK";
            this.btOK.UseVisualStyleBackColor = true;
            this.btOK.Click += new System.EventHandler(this.btOK_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.rbJustificationGiven);
            this.panel1.Controls.Add(this.rbAdminGrant);
            this.panel1.Controls.Add(this.rbEndOfRoad);
            this.panel1.Location = new System.Drawing.Point(48, 105);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(272, 69);
            this.panel1.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(381, 39);
            this.label1.TabIndex = 9;
            this.label1.Text = "Sorry, but this application won\'t work without access to your Microsoft calendar." +
    "\r\n\r\nPlease could you take a moment to confirm why authorisation wasn\'t given:-";
            // 
            // MsOauthConsent
            // 
            this.AcceptButton = this.btOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(412, 236);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btOK);
            this.Controls.Add(this.rbDunno);
            this.Controls.Add(this.rbDoubts);
            this.Controls.Add(this.rbBlocked);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MsOauthConsent";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Microsoft Authorisation Not Provided";
            this.TopMost = true;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.RadioButton rbBlocked;
        private System.Windows.Forms.RadioButton rbDoubts;
        private System.Windows.Forms.RadioButton rbJustificationGiven;
        private System.Windows.Forms.RadioButton rbAdminGrant;
        private System.Windows.Forms.RadioButton rbEndOfRoad;
        private System.Windows.Forms.RadioButton rbDunno;
        private System.Windows.Forms.Button btOK;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
    }
}