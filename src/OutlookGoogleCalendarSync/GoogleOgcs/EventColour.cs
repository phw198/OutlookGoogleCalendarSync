using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
//using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using System.Drawing;


namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class Palette {
        public String Id { get; }
        public String HexValue { get; }
        public Color RgbValue { get; }

        public Palette(String id, String hexValue, Color rgbValue) {
            this.Id = id;
            this.HexValue = hexValue;
            this.RgbValue = rgbValue;
        }
    }


    public class EventColour {
        private static readonly ILog log = LogManager.GetLogger(typeof(EventColour));        
        private List<Palette> colourPalette;

        public EventColour() { }

        public void Get() {
            log.Debug("Retrieving calendar Event colours.");
            Colors colours = null;
            colourPalette = new List<Palette>();
            try {
                colours = GoogleOgcs.Calendar.Instance.Service.Colors.Get().Execute();
            } catch (System.Exception ex) {
                log.Error("Failed retrieving calendar Event colours.");
                OGCSexception.Analyse(ex);
                return;
            }

            if (colours == null) log.Warn("No colours found!");
            else log.Debug(colours.Event__.Count() + " colours found.");
            
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Event__) {
                colourPalette.Add(new Palette(colour.Key, colour.Value.Background, OutlookOgcs.CategoryMap.RgbColour(colour.Value.Background)));
            }
        }

        //public static String HexValue(Color rgbColour) {
        //    String html = ColorTranslator.ToHtml(rgbColour);
        //    if (rgbColour.IsNamedColor)
        //        log.Fine(rgbColour.Name + " converted to " + html);
        //    else
        //        log.Fine(rgbColour.ToString() + " converted to " + html);
        //    return html;
        //}

        /// <summary>
        /// Find the closest colour offered by Google.
        /// </summary>
        /// <param name="colour">The colour to search with.</param>
        public String GetClosestColour(Color baseColour) {
            var colourDistance = colourPalette.Select(x => new { Value = x.Id, Diff = getDiff(x.RgbValue, baseColour) }).ToList();
            var minDistance = colourDistance.Min(x => x.Diff);
            return colourDistance.Find(x => x.Diff == minDistance).Value;
        }

        private int getDiff(Color colour, Color baseColour) {
            int a = colour.A - baseColour.A,
                r = colour.R - baseColour.R,
                g = colour.G - baseColour.G,
                b = colour.B - baseColour.B;
            return (a * a) + (r * r) + (g * g) + (b * b);
        }
    }
}
