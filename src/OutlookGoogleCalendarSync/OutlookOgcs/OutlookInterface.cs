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
        Folders Folders();
        Dictionary<string, MAPIFolder> CalendarFolders();
        MAPIFolder UseOutlookCalendar();
        void UseOutlookCalendar(MAPIFolder set);
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
        void GetAppointmentByID(String entryID, out AppointmentItem ai);
    }
}
