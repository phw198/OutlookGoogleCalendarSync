using System.Runtime.InteropServices;
using log4net;

namespace System.Windows.Forms {
    public static class OgcsMessageBox {
        private static readonly ILog log = LogManager.GetLogger(typeof(OgcsMessageBox));
        private static DialogResult dr;
                
        #region Window flashing
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO {
            public UInt32 cbSize;
            public IntPtr hwnd;
            public UInt32 dwFlags;
            public UInt32 uCount;
            public UInt32 dwTimeout;
        }

        [Flags]
        private enum flashMode {
            /// <summary>Stop flashing. The system restores the window to its original state.</summary>
            FLASHW_STOP = 0,
            /// <summary>Flash the window caption.</summary>
            FLASHW_CAPTION = 1,
            /// <summary>Flash the taskbar button.</summary>
            FLASHW_TRAY = 2,
            /// <summary>
            /// Flash both the window caption and taskbar button. 
            /// This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags.
            /// </summary>
            FLASHW_ALL = 3,
            /// <summary>Flash continuously, until the FLASHW_STOP flag is set.</summary>
            FLASHW_TIMER = 4,
            /// <summary>Flash continuously until the window comes to the foreground.</summary>
            FLASHW_TIMERNOFG = 12
        }

        /// <summary>
        /// Cause the window and taskbar icon to flash
        /// </summary>
        /// <param name="hWnd">The handle for the window to flash</param>
        /// <param name="fm">Bitwise flags</param>
        /// <returns></returns>
        private static bool flashWindow(IntPtr hWnd, flashMode fm) {
            FLASHWINFO fInfo = new FLASHWINFO();

            fInfo.cbSize = Convert.ToUInt32(Marshal.SizeOf(fInfo));
            fInfo.hwnd = hWnd;
            fInfo.dwFlags = (UInt32)fm;
            fInfo.uCount = UInt32.MaxValue;
            fInfo.dwTimeout = 0;

            return FlashWindowEx(ref fInfo);
        }
        #endregion

        /// <summary>
        /// Support cross-threading. Always show Main form and make it the owner of the MessageBox.
        /// </summary>
        /// <param name="text">Main text of the message</param>
        /// <param name="caption">Title of the box</param>
        /// <param name="buttons">Buttons to display</param>
        /// <param name="icon">Icon to display</param>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon) {
            OutlookGoogleCalendarSync.Forms.Main mainFrm = OutlookGoogleCalendarSync.Forms.Main.Instance;
            log.Debug(caption + ": " + text);

            if (mainFrm == null || mainFrm.IsDisposed)
                return MessageBox.Show(text, caption, buttons, icon);

            if (mainFrm.InvokeRequired) {
                mainFrm.Invoke(new System.Action(() => {
                    mainFrm.MainFormShow();
                    flashWindow(mainFrm.Handle, flashMode.FLASHW_ALL | flashMode.FLASHW_TIMERNOFG);
                    dr = MessageBox.Show(mainFrm, text, caption, buttons, icon);
                }));
            } else {
                mainFrm.MainFormShow();
                flashWindow(mainFrm.Handle, flashMode.FLASHW_ALL | flashMode.FLASHW_TIMERNOFG);
                dr = MessageBox.Show(mainFrm, text, caption, buttons, icon);
            }
            log.Debug("Response: " + dr.ToString());
            return dr;
        }

        /// <summary>
        /// Support cross-threading. Always show Main form and make it the owner of the MessageBox.
        /// </summary>
        /// <param name="text">Main text of the message</param>
        /// <param name="caption">Title of the box</param>
        /// <param name="buttons">Buttons to display</param>
        /// <param name="icon">Icon to display</param>
        /// <param name="defaultButton">Button to focus</param>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton) {
            OutlookGoogleCalendarSync.Forms.Main mainFrm = OutlookGoogleCalendarSync.Forms.Main.Instance;
            log.Debug(caption + ": " + text);

            if (mainFrm == null || mainFrm.IsDisposed)
                return MessageBox.Show(text, caption, buttons, icon, defaultButton);

            if (mainFrm.InvokeRequired) {
                mainFrm.Invoke(new System.Action(() => {
                    mainFrm.MainFormShow();
                    flashWindow(mainFrm.Handle, flashMode.FLASHW_ALL | flashMode.FLASHW_TIMERNOFG);
                    dr = MessageBox.Show(mainFrm, text, caption, buttons, icon, defaultButton);
                }));
            } else {
                mainFrm.MainFormShow();
                flashWindow(mainFrm.Handle, flashMode.FLASHW_ALL | flashMode.FLASHW_TIMERNOFG);
                dr = MessageBox.Show(mainFrm, text, caption, buttons, icon, defaultButton);
            }
            log.Debug("Response: " + dr.ToString());
            return dr;
        }
    }
}
