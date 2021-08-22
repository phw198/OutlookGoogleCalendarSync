using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class ErrorReporting : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorReporting));

        private String logFile;
        private static ErrorReporting instance;
        public static ErrorReporting Instance {
            get {
                if (instance == null || instance.IsDisposed) instance = new ErrorReporting();
                return instance;
            }
            set {
                instance = value;
            }
        }
            
        private ErrorReporting() {
            InitializeComponent();
            logFile = Path.Combine(log4net.GlobalContext.Properties["LogPath"].ToString(), log4net.GlobalContext.Properties["LogFilename"].ToString());
        }

        private void CloudLogging_Load(object sender, EventArgs e) {
            log.Debug("Asking user if they want to automatically report errors.");
            List<String> lines = new List<String>();
            using (FileStream logFileStream = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                StreamReader logFileReader = new StreamReader(logFileStream);
                while (!logFileReader.EndOfStream) {
                    lines.Add(logFileReader.ReadLine());
                }
            }
            foreach (String line in lines.Skip(lines.Count - 50).ToList()) {
                tbLog.Text += line + "\n";
            }
            tbLog.SelectionStart = tbLog.Text.Length;
            tbLog.ScrollToCaret();
        }

        private void btOpenLog_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start("explorer.exe", logFile);
        }

        private void CloudLogging_Shown(object sender, EventArgs e) {
            try {
                //Highlight the ERROR text and scroll so it's in view
                int lastError = tbLog.Text.LastIndexOf(" ERROR ") + 1;
                if (lastError == 0) lastError = tbLog.Text.LastIndexOf(" FATAL ") + 1;
                int highlightLength = tbLog.Text.Substring(lastError).IndexOf("\n");
                tbLog.Select(lastError, highlightLength);

                if (tbLog.SelectionStart != 0) {
                    tbLog.SelectionBackColor = System.Drawing.Color.Yellow;

                    int previousLineBreak = tbLog.Text.Substring(0, lastError).LastIndexOf("\n");
                    tbLog.SelectionStart = previousLineBreak;
                    tbLog.ScrollToCaret();
                }
                btYes.Focus();

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private void tbLog_Resize(object sender, EventArgs e) {
            tbLog.ScrollToCaret();
        }
    }
}
