using log4net;
using System;
using System.Reflection;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    class EventAttendee : Google.Apis.Calendar.v3.Data.EventAttendee {
        private static readonly ILog log = LogManager.GetLogger(typeof(EventAttendee));
        public const String EmailCloak = ".ogcs";

        private Google.Apis.Calendar.v3.Data.EventAttendee parent { get; set; }

        public EventAttendee() { }
        public EventAttendee(Google.Apis.Calendar.v3.Data.EventAttendee baseAttendee) {
            parent = baseAttendee;
            foreach (PropertyInfo prop in parent.GetType().GetProperties()) {
                try {
                    prop.SetValue(this, prop.GetValue(baseAttendee));
                } catch (System.Exception ex) {
                    log.Warn("Failed to set property " + prop.Name);
                    OGCSexception.Analyse(ex);
                }
            }
        }

        public new String Email {
            get { return decloakEmail(base.Email); }
            set { base.Email = cloakEmail(value); }
        }

        private static String decloakEmail(String email) {
            if (string.IsNullOrWhiteSpace(email)) return email;
            return email.TrimEnd(EmailCloak.ToCharArray());
        }
        private static String cloakEmail(String email) {
            if (string.IsNullOrWhiteSpace(email)) return email;
            String decloakedEmail = decloakEmail(email);
            return (decloakedEmail + (Settings.Instance.CloakEmail ? EmailCloak : ""));
        }
    }
}
