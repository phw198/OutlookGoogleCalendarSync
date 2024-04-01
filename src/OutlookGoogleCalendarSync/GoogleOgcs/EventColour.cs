using Google.Apis.Calendar.v3.Data;
using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;


namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class EventColour {
        public class Palette {
            private Boolean UseWebAppColours = true;
            public enum Type {
                Calendar,
                Event
            }
            private Type colourType = Type.Event;
            public String Id { get; internal set; }

            private String hexValue;
            public String HexValue {
                get {
                    if (UseWebAppColours) {
                        if (!string.IsNullOrEmpty(Id)) {
                            int idx = Convert.ToInt16(Id);
                            if (this.colourType == Type.Event) {
                                if (eventColourNames.ContainsKey(idx))
                                    return eventColourNames[idx].WebAppHexValue ?? hexValue;
                            } else {
                                if (calendarColourNames.ContainsKey(idx))
                                    return calendarColourNames[idx].WebAppHexValue ?? hexValue;
                            }
                        } else
                            return hexValue;
                    }
                    return hexValue;
                }
                internal set { hexValue = value; }
            }

            private Color rgbValue;
            public Color RgbValue {
                get {
                    if (UseWebAppColours) {
                        if (rgbConvertedFromHex.IsEmpty && HexValue != null) rgbConvertedFromHex = OutlookOgcs.Categories.Map.RgbColour(HexValue);
                        return rgbConvertedFromHex;
                    }
                    return rgbValue;
                }
                internal set { rgbValue = value; }
            }

            private Color rgbConvertedFromHex;

            public String Name {
                get {
                    String name = "";
                    try {
                        name = GetColourName(Id);
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse(ex);
                        name = HexValue;
                    }
                    return name;
                }
            }

            public static Palette NullPalette = new Palette(Type.Event, null, null, Color.Transparent);

            public Palette(Type colourType, String id, String hexValue, Color rgbValue) {
                this.colourType = colourType;
                this.Id = id;
                this.HexValue = hexValue;
                this.RgbValue = rgbValue;
            }

            public override String ToString() {
                return "Type: " + this.colourType + "; ID: " + Id + "; HexValue: " + HexValue + "; RgbValue: " + RgbValue + "; Name: " + Name;
            }

            private class Metadata {
                internal String Name { get; }
                internal String WebAppHexValue { get; }

                internal Metadata(String name, String webAppHexValue) {
                    Name = name;
                    WebAppHexValue = webAppHexValue;
                }
            }

            private static Dictionary<int, Metadata> calendarColourNames = new Dictionary<int, Metadata> {
                { 0, new Metadata("Custom colour", null) }, //Although custom colour has it's own (correct!) Hex code, Google sets the ID to closest match!
                { 1, new Metadata("Cocoa", "#795548") },
                { 2, new Metadata("Flamingo", "#E67C73") },
                { 3, new Metadata("Tomato", "#D50000") },
                { 4, new Metadata("Tangerine", "#F4511E") },
                { 5, new Metadata("Pumpkin", "#EF6C00") },
                { 6, new Metadata("Mango", "#F09300") },
                { 7, new Metadata("Eucalyptus", "#009688") },
                { 8, new Metadata("Basil", "#0B8043") },
                { 9, new Metadata("Pistachio", "#7CB342") },
                { 10, new Metadata("Avocado", "#C0CA33") },
                { 11, new Metadata("Citron", "#E4C441") },
                { 12, new Metadata("Banana", "#F6BF26") },
                { 13, new Metadata("Sage", "#33B679") },
                { 14, new Metadata("Peacock", "#039BE5") },
                { 15, new Metadata("Cobalt", "#4285F4") },
                { 16, new Metadata("Blueberry", "#3F51B5") },
                { 17, new Metadata("Lavendar", "#7986CB") },
                { 18, new Metadata("Wisteria", "#B39DDB") },
                { 19, new Metadata("Graphite", "#616161") },
                { 20, new Metadata("Birch", "#A79B8E") },
                { 21, new Metadata("Beetroot", "#AD1457") },
                { 22, new Metadata("Cherry Blossom", "#D81B60") },
                { 23, new Metadata("Grape", "#8E24AA") },
                { 24, new Metadata("Amethyst", "#9E69AF") }
            };

            private static Dictionary<int, Metadata> eventColourNames = new Dictionary<int, Metadata> {
                { 0, new Metadata("Calendar Default", null) },
                { 1, new Metadata("Lavendar", "#7986CB") },
                { 2, new Metadata("Sage", "#33B679") },
                { 3, new Metadata("Grape", "#8E24AA") },
                { 4, new Metadata("Flamingo", "#E67C73") },
                { 5, new Metadata("Banana", "#F6BF26") },
                { 6, new Metadata("Tangerine", "#F4511E") },
                { 7, new Metadata("Peacock", "#039BE5") },
                { 8, new Metadata("Graphite", "#616161") },
                { 9, new Metadata("Blueberry", "#3F51B5") },
                { 10, new Metadata("Basil", "#0B8043") },
                { 11, new Metadata("Tomato", "#D50000") }
            };

            /// <summary>
            /// Get colour ID from the name
            /// </summary>
            /// <param name="name">The name of the colour</param>
            /// <returns>The ID number</returns>
            public static String GetColourId(String name) {
                String id = null;
                try {
                    id = eventColourNames.First(n => (n.Value as Metadata).Name == name).Key.ToString();
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Could not find colour ID for '" + name + "'.", ex);
                }
                return id;
            }

            /// <summary>
            /// Get colour name from the ID
            /// </summary>
            /// <param name="id">The colour ID</param>
            /// <returns>The colour name</returns>
            public static String GetColourName(String id) {
                if (string.IsNullOrEmpty(id)) return null;

                String name = null;
                try {
                    int idx = Convert.ToInt16(id);
                    if (eventColourNames.ContainsKey(idx))
                        name = (eventColourNames[idx] as Metadata).Name;
                    else
                        log.Error("GetColourName(): ID '" + id + "' not found.");
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Could not find colour name for '" + id + "'.", ex);
                }
                return name;
            }
        }

        private static readonly ILog log = LogManager.GetLogger(typeof(EventColour));
        private List<Palette> calendarPalette;
        private List<Palette> eventPalette;
        public EventColour() { }

        /// <summary>
        /// All event colours, including currently used calendar "custom" colour
        /// </summary>
        public List<Palette> ActivePalette {
            get {
                SettingsStore.Calendar profile = Settings.Profile.InPlay();
                List<Palette> activePalette = new List<Palette>();
                if (profile.UseGoogleCalendar?.Id == null) return activePalette;

                if (profile.UseGoogleCalendar.ColourId == "0") {
                    GoogleOgcs.Calendar.Instance.GetCalendars();
                    profile.UseGoogleCalendar.ColourId = GoogleOgcs.Calendar.Instance.CalendarList.Find(c => c.Id == profile.UseGoogleCalendar.Id)?.ColourId ?? "0";
                }

                //Palette currentCal = calendarPalette.Find(p => p.Id == profile.UseGoogleCalendar.ColourId);
                Palette currentCal = null;
                foreach (Palette cal in calendarPalette) {
                    if (cal.Id == profile.UseGoogleCalendar.ColourId) {
                        currentCal = cal;
                        break;
                    }
                }
                if (currentCal != null)
                    activePalette.Add(new Palette(Palette.Type.Calendar, "0", currentCal.HexValue, currentCal.RgbValue));

                activePalette.AddRange(eventPalette);
                return activePalette;
            }
        }

        public Boolean IsCached() {
            return (calendarPalette != null && calendarPalette.Count != 0 && eventPalette != null && eventPalette.Count != 0);
        }

        /// <summary>
        /// Retrieve calendar's Event colours from Google
        /// </summary>
        public void Get() {
            log.Debug("Retrieving calendar Event colours.");
            Colors colours = null;
            calendarPalette = new List<Palette>();
            eventPalette = new List<Palette>();
            int backoff = 0;
            try {
                while (backoff < GoogleOgcs.Calendar.BackoffLimit) {
                    try {
                        colours = GoogleOgcs.Calendar.Instance.Service.Colors.Get().Execute();
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (GoogleOgcs.Calendar.HandleAPIlimits(ref ex, null)) {
                            case GoogleOgcs.Calendar.ApiException.throwException: throw;
                            case GoogleOgcs.Calendar.ApiException.freeAPIexhausted:
                                OGCSexception.LogAsFail(ref ex);
                                OGCSexception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(GoogleOgcs.Calendar.Instance.SubscriptionInvite, ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case GoogleOgcs.Calendar.ApiException.backoffThenRetry:
                                backoff++;
                                if (backoff == GoogleOgcs.Calendar.BackoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve Event colours failed.");
                                    throw;
                                } else {
                                    int backoffDelay = (int)Math.Pow(2, backoff);
                                    log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoffDelay * 1000);
                                }
                                break;
                        }
                    }
                }
            } catch (System.Exception ex) {
                log.Error("Failed retrieving calendar Event colours.");
                OGCSexception.Analyse(ex);
                throw;
            }

            if (colours == null) log.Warn("No colours found!");
            else log.Debug(colours.Event__.Count() + " event colours and " + colours.Calendar.Count() + " calendars (with a colour) found.");

            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Event__) {
                eventPalette.Add(new Palette(Palette.Type.Event, colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Calendar) {
                calendarPalette.Add(new Palette(Palette.Type.Calendar, colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
            Forms.Main.Instance.miColourBuildPicker_Click(null, null);
        }
        
        /// <summary>
        /// Build colour list from those downloaded from Google.
        /// </summary>
        /// <param name="clb">The checklistbox to populate with the colours.</param>
        public void BuildPicker(System.Windows.Forms.CheckedListBox clb) {
            clb.BeginUpdate();
            clb.Items.Clear();
            clb.Items.Add("<Default calendar colour>");
            foreach (Palette colour in GoogleOgcs.Calendar.Instance.ColourPalette.eventPalette) {
                clb.Items.Add(colour.Name);
            }
            foreach (String colour in Forms.Main.Instance.ActiveCalendarProfile.Colours) {
                try {
                    clb.SetItemChecked(clb.Items.IndexOf(colour), true);
                } catch { /* Colour "colour" no longer exists */ }
            }
            clb.EndUpdate();
        }

        /// <summary>
        /// Get the Google Palette from its Google ID
        /// </summary>
        /// <param name="colourId">Google ID</param>
        public Palette GetColour(String colourId) {
            Palette gColour = this.ActivePalette.Where(x => x.Id == colourId).FirstOrDefault();
            if (colourId != "0" && gColour != null)
                return gColour;
            else
                return Palette.NullPalette;
        }

        /// <summary>
        /// Find the closest colour palette offered by Google.
        /// </summary>
        /// <param name="colour">The colour to search with.</param>
        public Palette GetClosestColour(Color baseColour) {
            try {
                var colourDistance = ActivePalette.Select(x => new { Value = x, Diff = GetDiff(x.RgbValue, baseColour) }).ToList();
                var minDistance = colourDistance.Min(x => x.Diff);
                return colourDistance.Find(x => x.Diff == minDistance).Value;
            } catch (System.Exception ex) {
                log.Warn("Failed to get closest Event colour for " + baseColour.Name);
                OGCSexception.Analyse(ex);
                return Palette.NullPalette;
            }
        }

        public static int GetDiff(Color colour, Color baseColour) {
            int a = colour.A - baseColour.A,
                r = colour.R - baseColour.R,
                g = colour.G - baseColour.G,
                b = colour.B - baseColour.B;
            return (a * a) + (r * r) + (g * g) + (b * b);
        }
    }
}
