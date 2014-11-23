using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleCalendarSync {
    public interface OutlookInterface {
        void Connect();
        List<String> Accounts();
        Dictionary<string, MAPIFolder> CalendarFolders();
        MAPIFolder UseOutlookCalendar();
        void UseOutlookCalendar(MAPIFolder set);
        String CurrentUserSMTP();
        String CurrentUserName();

        void CreateCalendarEntries(List<Event> events);
        void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated);

        String GetRecipientEmail(Recipient recipient);
        Event AddGoogleAttendee(EventAttendee ea, Event ev);
        Boolean CompareRecipientsToAttendees(AppointmentItem ai, Event ev, Dictionary<String,Boolean> attendeesFromDescription, StringBuilder sb, ref int itemModified);
        
    }
}
