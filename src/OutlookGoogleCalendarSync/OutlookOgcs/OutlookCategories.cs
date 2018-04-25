using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Outlook = Microsoft.Office.Interop.Outlook;
using log4net;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public class Categories {
        private static readonly ILog log = LogManager.GetLogger(typeof(Categories));
        private Outlook.Categories categories;
        public String Delimiter { get; }

        public Categories() {
            try {
                Delimiter = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator + " ";
            } catch (System.Exception ex) {
                log.Error("Failed to get system ListSeparator value.");
                OGCSexception.Analyse(ex);
                Delimiter = ", ";
            }
        }
        
        public void Get(Outlook.Application oApp, Outlook.Store store) {
            if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.DefaultMailbox)
                this.categories = oApp.Session.Categories;
            else {
                try {
                    this.categories = store.GetType().GetProperty("Categories").GetValue(store, null) as Outlook.Categories;
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex, true);
                    this.categories = oApp.Session.Categories;
                }
            }
        }

        //private static Dictionary<Outlook.OlCategoryColor, Dictionary<String, String>> Index;

        //public void BuildIndex() {
        //    foreach (Outlook.Category category in Outlook.Categ this.categories) {
        //        CategoryMap map = new CategoryMap( {
        //            Name = category.Name,
        //            OutlookColourNumber = category.Color);
        //        log.Debug(category.Color.ToString());
        //    }
        //}

        /// <summary>
        /// Get the Outlook category type from the name given to the category
        /// </summary>
        /// <param name="categoryName">The user named Outlook category</param>
        /// <returns>The Outlook category type</returns>
        public Outlook.OlCategoryColor? OutlookColour(String categoryName) {
            if (string.IsNullOrEmpty(categoryName)) log.Warn("Category name is empty.");

            foreach (Outlook.Category category in this.categories) {
                if (category.Name == categoryName.Trim()) return category.Color;
            }
            return null;
        }

        //public void List() {
        //    foreach (Outlook.Category category in this.categories) {
        //        log.Debug(category.Name);
        //        log.Debug(category.Color.ToString());
        //    }
        //}

        /// <summary>
        /// Return all the category names as a list of strings.
        /// </summary>
        public List<String> GetNames() {
            List<String> names = new List<String>();
            foreach (Outlook.Category category in this.categories) {
                names.Add(category.Name);
            }
            return names;
        }
    }

    public class CategoryMap {
        private static readonly ILog log = LogManager.GetLogger(typeof(CategoryMap));

        //Source: https://msdn.microsoft.com/en-us/library/ee203806%28v=exchg.80%29.aspx
        public static Dictionary<Outlook.OlCategoryColor, Color> Colours { get; }
        static CategoryMap() {
            Colours = new Dictionary<Outlook.OlCategoryColor, Color> {
                { Outlook.OlCategoryColor.olCategoryColorBlack, Color.FromArgb(28,28,28) },
                { Outlook.OlCategoryColor.olCategoryColorBlue, Color.FromArgb(50, 103, 184) },
                { Outlook.OlCategoryColor.olCategoryColorDarkBlue, Color.FromArgb(42, 81, 145 ) },
                { Outlook.OlCategoryColor.olCategoryColorDarkGray, Color.FromArgb(165, 165, 165) },
                { Outlook.OlCategoryColor.olCategoryColorDarkGreen, Color.FromArgb(53, 121, 43) },
                { Outlook.OlCategoryColor.olCategoryColorDarkMaroon, Color.FromArgb(130, 55, 95) },
                { Outlook.OlCategoryColor.olCategoryColorDarkOlive, Color.FromArgb(95, 108, 58) },
                { Outlook.OlCategoryColor.olCategoryColorDarkOrange, Color.FromArgb(177, 79, 13 ) },
                { Outlook.OlCategoryColor.olCategoryColorDarkPeach, Color.FromArgb(171, 123, 5 ) },
                { Outlook.OlCategoryColor.olCategoryColorDarkPurple, Color.FromArgb(80, 50, 143) },
                { Outlook.OlCategoryColor.olCategoryColorDarkRed, Color.FromArgb(175, 30, 37) },
                { Outlook.OlCategoryColor.olCategoryColorDarkSteel, Color.FromArgb(140, 156, 189) },
                { Outlook.OlCategoryColor.olCategoryColorDarkTeal, Color.FromArgb(46, 125, 100) },
                { Outlook.OlCategoryColor.olCategoryColorDarkYellow, Color.FromArgb(153, 148, 0) },
                { Outlook.OlCategoryColor.olCategoryColorGray, Color.FromArgb(196, 196, 196) },
                { Outlook.OlCategoryColor.olCategoryColorGreen, Color.FromArgb(74, 182, 63) },
                { Outlook.OlCategoryColor.olCategoryColorMaroon, Color.FromArgb(163, 78, 120) },
                { Outlook.OlCategoryColor.olCategoryColorNone, Color.FromArgb(255, 255, 255) },
                { Outlook.OlCategoryColor.olCategoryColorOlive, Color.FromArgb(133, 154, 82) },
                { Outlook.OlCategoryColor.olCategoryColorOrange, Color.FromArgb(240, 108, 21) },
                { Outlook.OlCategoryColor.olCategoryColorPeach, Color.FromArgb(255, 202, 76) },
                { Outlook.OlCategoryColor.olCategoryColorPurple, Color.FromArgb(97, 61, 180) },
                { Outlook.OlCategoryColor.olCategoryColorRed, Color.FromArgb(214, 37, 46) },
                { Outlook.OlCategoryColor.olCategoryColorSteel, Color.FromArgb(196, 204, 221) },
                { Outlook.OlCategoryColor.olCategoryColorTeal, Color.FromArgb(64, 189, 149) },
                { Outlook.OlCategoryColor.olCategoryColorYellow, Color.FromArgb(255, 254, 61) }
            };
        }

        /// <summary>
        /// Convert from Outlook category colour to Color
        /// </summary>
        public static Color RgbColour(Outlook.OlCategoryColor colour) {
            log.Fine("Converting " + colour + " to RGB value.");
            return Colours[colour];
        }

        /// <summary>
        /// Convert from HTML hex string to Color
        /// </summary>
        public static Color RgbColour(String hexColour) {
            Color colour = ColorTranslator.FromHtml(hexColour);
            log.Fine("Converted " + hexColour + " to " + colour.ToString());
            return colour;
        }

        //public String Name { get; set; }
        //public Int16 OutlookColourNumber { get; set; }
        //public String GoogleColourHex { get; set; }
    }
}
