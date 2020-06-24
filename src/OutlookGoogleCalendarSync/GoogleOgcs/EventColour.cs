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
            public String Id { get; }

            private String hexValue;
            public String HexValue {
                get {
                    if (UseWebAppColours) {
                        if (!string.IsNullOrEmpty(Id)) {
                            int idx = Convert.ToInt16(Id);
                            if (names.ContainsKey(idx))
                                return names[idx].WebAppHexValue ?? hexValue;
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
                    if (UseWebAppColours)
                        return OutlookOgcs.Categories.Map.RgbColour(HexValue);
                    return rgbValue;
                }
                internal set { rgbValue = value; }
            }

            public String Name { get {
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

            public static Palette NullPalette = new Palette(null, null, Color.Transparent);

            public Palette(String id, String hexValue, Color rgbValue) {
                this.Id = id;
                this.HexValue = hexValue;
                this.RgbValue = rgbValue;
            }

            public override String ToString() {
                return "ID: " + Id + "; HexValue: " + HexValue + "; RgbValue: " + RgbValue +"; Name: "+ Name;
            }
            
            private class Metadata {
                internal String Name { get; }
                internal String WebAppHexValue { get; }
                
                internal Metadata(String name, String webAppHexValue) {
                    Name = name;
                    WebAppHexValue = webAppHexValue;
                }
            }

            private static Dictionary<int, Metadata> names = new Dictionary<int, Metadata> {
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
                    id = names.First(n => (n.Value as Metadata).Name == name).Key.ToString();
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
                    if (names.ContainsKey(idx))
                        name = (names[idx] as Metadata).Name;
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
        /// <summary>
        /// All event colours, including currently used calendar "custom" colour
        /// </summary>
        public List<Palette> ActivePalette {
            get {
                List<Palette> activePalette = new List<Palette>();
                if (Settings.Instance.UseGoogleCalendar == null) return activePalette;

                //Palette currentCal = calendarPalette.Find(p => p.Id == Settings.Instance.UseGoogleCalendar.ColourId);
                Palette currentCal = null;
                foreach (Palette cal in calendarPalette) {
                    if (cal.Id == Settings.Instance.UseGoogleCalendar.ColourId) {
                        currentCal = cal;
                        break;
                    }
                }
                if (currentCal != null)
                    activePalette.Add(new Palette("0", currentCal.HexValue, currentCal.RgbValue));

                activePalette.AddRange(eventPalette);
                return activePalette;
            }
        }

        public EventColour() { }

        /// <summary>
        /// Retrieve calendar's Event colours from Google
        /// </summary>
        public void Get() {
            log.Debug("Retrieving calendar Event colours.");
            Colors colours = null;
            calendarPalette = new List<Palette>();
            eventPalette = new List<Palette>();
            try {
                colours = GoogleOgcs.Calendar.Instance.Service.Colors.Get().Execute();
            } catch (System.Exception ex) {
                log.Error("Failed retrieving calendar Event colours.");
                OGCSexception.Analyse(ex);
                return;
            }

            if (colours == null) log.Warn("No colours found!");
            else log.Debug(colours.Event__.Count() + " event colours and "+ colours.Calendar.Count() +" calendars (with a colour) found.");
            
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Event__) {
                eventPalette.Add(new Palette(colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Calendar) {
                calendarPalette.Add(new Palette(colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
        }

        /// <summary>
        /// Get the Google Palette from its Google ID
        /// </summary>
        /// <param name="colourId">Google ID</param>
        public Palette GetColour(String colourId) {
            Palette gColour = this.ActivePalette.Where(x => x.Id == colourId).FirstOrDefault();
            if (gColour != null)
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
