namespace System.Windows.Forms {
    public static class OgcsMessageBox {
        private static DialogResult dr;
                
        /// <summary>
        /// Support cross-threading. Always show Main form and make it the owner of the MessageBox.
        /// </summary>
        /// <param name="text">Main text of the message</param>
        /// <param name="caption">Title of the box</param>
        /// <param name="buttons">Buttons to display</param>
        /// <param name="icon">Icon to display</param>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon) {
            OutlookGoogleCalendarSync.Forms.Main mainFrm = OutlookGoogleCalendarSync.Forms.Main.Instance;

            if (mainFrm == null || mainFrm.IsDisposed)
                return MessageBox.Show(text, caption, buttons, icon);

            if (mainFrm.InvokeRequired) {
                mainFrm.Invoke(new System.Action(() => {
                    mainFrm.MainFormShow();
                    dr = MessageBox.Show(mainFrm, text, caption, buttons, icon);
                }));
            } else {
                mainFrm.MainFormShow();
                dr = MessageBox.Show(mainFrm, text, caption, buttons, icon);
            }
            return dr;
        }
    }
}
