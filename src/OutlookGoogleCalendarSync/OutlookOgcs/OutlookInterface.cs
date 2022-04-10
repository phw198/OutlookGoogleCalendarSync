using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public interface Interface {
        void Connect();
        void Disconnect(Boolean onlyWhenNoGUI = false);
        Boolean NoGUIexists();
        Folders Folders();
        Dictionary<string, MAPIFolder> CalendarFolders();
        MAPIFolder UseOutlookCalendar();
        NameSpace GetCurrentUser(NameSpace oNS);
        String CurrentUserSMTP();
        String CurrentUserName();
        Boolean Offline();
        void RefreshCategories();

        String GetRecipientEmail(Recipient recipient);
        OlExchangeConnectionMode ExchangeConnectionMode();
        String GetGlobalApptID(AppointmentItem ai);

        Event IANAtimezone_set(Event ev, AppointmentItem ai);
        void WindowsTimeZone_get(AppointmentItem ai, out String startTz, out String endTz);
        AppointmentItem WindowsTimeZone_set(AppointmentItem ai, Event ev, String attr = "Both", Boolean onlyTZattribute = false);
        /// <summary>
        /// Filter Outlook Item according to specified filter.
        /// </summary>
        /// <param name="outlookItems">Items to be filtered</param>
        /// <param name="filter">The logic by which to perform filter</param>
        /// <returns>Filtered items</returns>
        List<Object> FilterItems(Items outlookItems, String filter);
        MAPIFolder GetFolderByID(String entryID);
        void GetAppointmentByID(String entryID, out AppointmentItem ai);
    }
}
