using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class CloudLogging : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(CloudLogging));

        private String logFile;

        public CloudLogging() {
            InitializeComponent();
            logFile = Path.Combine(log4net.GlobalContext.Properties["LogPath"].ToString(), log4net.GlobalContext.Properties["LogFilename"].ToString());
        }

        private void CloudLogging_Load(object sender, EventArgs e) {
            log.Debug("Asking user if they want to upload errors to Google Stackdriver Logging.");
            List<String> lines = new List<String>();
            using (FileStream logFileStream = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                StreamReader logFileReader = new StreamReader(logFileStream);
                while (!logFileReader.EndOfStream) {
                    lines.Add(logFileReader.ReadLine());
                }
            }
            foreach (String line in lines.Skip(lines.Count - 50).ToList()) {
                tbLog.Text += line +"\r\n";
            }

            tbLog.Text = tbLog.Text.TrimEnd(new char[] { '\r', '\n' });
            tbLog.SelectionStart = tbLog.Text.Length;
            tbLog.ScrollToCaret();
            System.Windows.Forms.Application.DoEvents();
        }

        private void btOpenLog_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start(logFile);
        }
    }
}
