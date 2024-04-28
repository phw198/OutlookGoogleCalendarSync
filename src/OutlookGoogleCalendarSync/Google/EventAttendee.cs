using log4net;
using System;
using System.Reflection;

namespace OutlookGoogleCalendarSync.Google {
    class EventAttendee : global::Google.Apis.Calendar.v3.Data.EventAttendee {
        private static readonly ILog log = LogManager.GetLogger(typeof(EventAttendee));
        public const String EmailCloak = ".ogcs";

        private global::Google.Apis.Calendar.v3.Data.EventAttendee parent { get; set; }

        public EventAttendee() { }
        public EventAttendee(global::Google.Apis.Calendar.v3.Data.EventAttendee baseAttendee) {
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
            set { base.Email = CloakEmail(value); }
        }

        private static String decloakEmail(String email) {
            if (string.IsNullOrWhiteSpace(email)) return email;
            if (email.EndsWith(EmailCloak)) return email.Substring(0, email.Length - EmailCloak.Length);
            return email;
        }
        public static String CloakEmail(String email) {
            if (string.IsNullOrWhiteSpace(email)) return email;
            String decloakedEmail = decloakEmail(email);
            return (decloakedEmail + (Sync.Engine.Calendar.Instance.Profile.CloakEmail ? EmailCloak : ""));
        }

        public Boolean IsCloaked() {
            return (base.Email.EndsWith(EmailCloak));
        }
    }
}
