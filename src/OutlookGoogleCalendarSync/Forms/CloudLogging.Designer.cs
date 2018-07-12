namespace OutlookGoogleCalendarSync.Forms {
    partial class CloudLogging {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CloudLogging));
            this.btYes = new System.Windows.Forms.Button();
            this.tbPanel = new System.Windows.Forms.Panel();
            this.tbLog = new System.Windows.Forms.TextBox();
            this.btNo = new System.Windows.Forms.Button();
            this.lTitle = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btOpenLog = new System.Windows.Forms.Button();
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbNotes = new System.Windows.Forms.TextBox();
            this.tbPanel.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btYes
            // 
            this.btYes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btYes.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.btYes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btYes.Location = new System.Drawing.Point(600, 378);
            this.btYes.Name = "btYes";
            this.btYes.Size = new System.Drawing.Size(75, 23);
            this.btYes.TabIndex = 1;
            this.btYes.Text = "Yes";
            this.btYes.UseVisualStyleBackColor = true;
            // 
            // tbPanel
            // 
            this.tbPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbPanel.Controls.Add(this.tbLog);
            this.tbPanel.Location = new System.Drawing.Point(15, 216);
            this.tbPanel.Name = "tbPanel";
            this.tbPanel.Size = new System.Drawing.Size(660, 150);
            this.tbPanel.TabIndex = 2;
            // 
            // tbLog
            // 
            this.tbLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbLog.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbLog.Location = new System.Drawing.Point(0, 0);
            this.tbLog.Multiline = true;
            this.tbLog.Name = "tbLog";
            this.tbLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbLog.Size = new System.Drawing.Size(660, 150);
            this.tbLog.TabIndex = 0;
            this.tbLog.WordWrap = false;
            // 
            // btNo
            // 
            this.btNo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btNo.DialogResult = System.Windows.Forms.DialogResult.No;
            this.btNo.Location = new System.Drawing.Point(519, 378);
            this.btNo.Name = "btNo";
            this.btNo.Size = new System.Drawing.Size(75, 23);
            this.btNo.TabIndex = 3;
            this.btNo.Text = "No";
            this.btNo.UseVisualStyleBackColor = true;
            // 
            // lTitle
            // 
            this.lTitle.AutoSize = true;
            this.lTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lTitle.Location = new System.Drawing.Point(12, 9);
            this.lTitle.Name = "lTitle";
            this.lTitle.Size = new System.Drawing.Size(303, 18);
            this.lTitle.TabIndex = 5;
            this.lTitle.Text = "Send error details to OGCS developer?";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 200);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Error Report:";
            // 
            // btOpenLog
            // 
            this.btOpenLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btOpenLog.Location = new System.Drawing.Point(16, 378);
            this.btOpenLog.Name = "btOpenLog";
            this.btOpenLog.Size = new System.Drawing.Size(104, 23);
            this.btOpenLog.TabIndex = 7;
            this.btOpenLog.Text = "Open Full Log";
            this.btOpenLog.UseVisualStyleBackColor = true;
            this.btOpenLog.Click += new System.EventHandler(this.btOpenLog_Click);
            // 
            // txtInfo
            // 
            this.txtInfo.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo.Location = new System.Drawing.Point(15, 39);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(660, 146);
            this.txtInfo.TabIndex = 8;
            this.txtInfo.Text = resources.GetString("txtInfo.Text");
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbNotes);
            this.groupBox1.Location = new System.Drawing.Point(16, 87);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(658, 68);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Notes";
            // 
            // tbNotes
            // 
            this.tbNotes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbNotes.BackColor = System.Drawing.SystemColors.Control;
            this.tbNotes.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbNotes.Location = new System.Drawing.Point(17, 20);
            this.tbNotes.Multiline = true;
            this.tbNotes.Name = "tbNotes";
            this.tbNotes.Size = new System.Drawing.Size(635, 42);
            this.tbNotes.TabIndex = 0;
            this.tbNotes.Text = resources.GetString("tbNotes.Text");
            // 
            // CloudLogging
            // 
            this.AcceptButton = this.btYes;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btNo;
            this.ClientSize = new System.Drawing.Size(687, 413);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btOpenLog);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lTitle);
            this.Controls.Add(this.btNo);
            this.Controls.Add(this.tbPanel);
            this.Controls.Add(this.btYes);
            this.Controls.Add(this.txtInfo);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(703, 431);
            this.Name = "CloudLogging";
            this.Text = "OGCS Error Encountered";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.CloudLogging_Load);
            this.tbPanel.ResumeLayout(false);
            this.tbPanel.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btYes;
        private System.Windows.Forms.Panel tbPanel;
        private System.Windows.Forms.Button btNo;
        private System.Windows.Forms.Label lTitle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btOpenLog;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.TextBox tbLog;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbNotes;
    }
}