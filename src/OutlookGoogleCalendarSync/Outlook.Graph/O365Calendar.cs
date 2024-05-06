using Ogcs = OutlookGoogleCalendarSync;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.Graph;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        private static Dictionary<String, OutlookCalendarListEntry> calendarFolders = new Dictionary<string, OutlookCalendarListEntry>();
        public static Dictionary<String, OutlookCalendarListEntry> CalendarFolders {
            get { return calendarFolders; }
        }

        /// <summary>Retrieve calendar list from the cloud.</summary>
        public static Dictionary<String, OutlookCalendarListEntry> GetCalendars() {
            calendarFolders = new();
            List<Microsoft.Graph.Calendar> cals = new();

            var graphThread = new System.Threading.Thread(() => {
                try {
                    Microsoft.Graph.IUserCalendarsCollectionPage calPage = Outlook.Calendar.Instance.GraphClient.Me.Calendars.Request().GetAsync().Result;
                    cals.AddRange(calPage.CurrentPage);
                    while (calPage.NextPageRequest != null) {
                        calPage = calPage.NextPageRequest.GetAsync().Result;
                        cals.AddRange(calPage.CurrentPage);
                    }
                } catch (System.Exception ex) {
                    log.Debug(ex.ToString());
                }
            });
            graphThread.Start();
            while (graphThread.IsAlive) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(250);
            }

            foreach (Microsoft.Graph.Calendar cal in cals) {
                if (cal.AdditionalData.ContainsKey("isDefaultCalendar") && (Boolean)cal.AdditionalData["isDefaultCalendar"])
                    cal.Name = "Default " + cal.Name;
                log.Debug(cal.Name);
                calendarFolders.Add(cal.Name, new OutlookCalendarListEntry(cal));
            }

            return calendarFolders;
        }
    }
}
