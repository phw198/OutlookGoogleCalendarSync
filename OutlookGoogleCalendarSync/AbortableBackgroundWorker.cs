using System.ComponentModel;
using System.Threading;
using log4net;

namespace OutlookGoogleCalendarSync {
    public class AbortableBackgroundWorker : BackgroundWorker {
        private static readonly ILog log = LogManager.GetLogger(typeof(AbortableBackgroundWorker));

        private Thread workerThread;

        protected override void OnDoWork(DoWorkEventArgs e) {
            workerThread = Thread.CurrentThread;
            try {
                base.OnDoWork(e);
            } catch (ThreadAbortException ex) {
                log.Error(ex.Message);
                log.Error(ex.StackTrace);
                e.Cancel = true; //We must set Cancel property to true!
                Thread.ResetAbort(); //Prevents ThreadAbortException propagation
            }
        }

        public void Abort() {
            if (workerThread != null) {
                log.Info("Aborting background thread...");
                
                workerThread.Interrupt();
                workerThread.Abort();
                workerThread = null;
                log.Info("Aborted.");
            }
        }
    }
}
