using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

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
        
        /// <summary>
        /// Get the categories as configured in Outlook and store in class
        /// </summary>
        /// <param name="oApp"></param>
        /// <param name="store"></param>
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

        /// <summary>
        /// Get the Outlook category from the name given to the category
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

        /// <summary>
        /// Get the Outlook categorys as List of ColourInfo
        /// </summary>
        /// <returns>List to be used in dropdown, for example</returns>
        public List<Extensions.ColourPicker.ColourInfo> DropdownItems() {
            List<Extensions.ColourPicker.ColourInfo> items = new List<Extensions.ColourPicker.ColourInfo>();
            foreach (Outlook.Category category in this.categories) {
                items.Add(new Extensions.ColourPicker.ColourInfo(category.Color, CategoryMap.RgbColour(category.Color), category.Name));
            }
            return items;
        }

        /// <summary>
        /// Get the Outlook category name for a given category colour.
        /// If category not yet used, new one added of the form "OGCS [colour]"
        /// </summary>
        /// <param name="olCategory">The Outlook category to search by</param>
        /// <param name="categoryName">Optional: The Outlook category name to also search by</param>
        /// <returns>The matching category name</returns>
        public String FindName(Outlook.OlCategoryColor olCategory, String categoryName = null) {
            if (olCategory == Outlook.OlCategoryColor.olCategoryColorNone) return "";

            Outlook.Category failSafeCategory = null;
            foreach (Outlook.Category category in this.categories) {
                if (category.Color == olCategory) {
                    if (categoryName == null) {
                        if (category.Name.StartsWith("OGCS ")) return category.Name;
                    } else {
                        if (category.Name == categoryName) return category.Name;
                        if (category.Name.StartsWith("OGCS ")) failSafeCategory = category;
                    }
                }
            }

            if (failSafeCategory != null) {
                log.Warn("Failed to find Outlook category " + olCategory.ToString() + " with name '" + categoryName + "'");
                log.Debug("Using category with name \"" + failSafeCategory.Name + "\" instead.");
                return failSafeCategory.Name;
            }

            log.Debug("Did not find Outlook category " + olCategory.ToString() + (categoryName == null ? "" : " \"" + categoryName + "\""));
            Outlook.Category newCategory = categories.Add("OGCS " + FriendlyCategoryName(olCategory), olCategory);
            log.Info("Added new Outlook category \"" + newCategory.Name + "\" for " + newCategory.Color.ToString());
            return newCategory.Name;
        }

        public static String FriendlyCategoryName(Outlook.OlCategoryColor olCategory) {
            return olCategory.ToString().Replace("olCategoryColor", "").Replace("Dark", "Dark ");
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

        public static Outlook.OlCategoryColor GetClosestCategory(GoogleOgcs.Palette basePalette) {
            try {
                var colourDistance = Colours.Select(x => new { Value = x, Diff = GoogleOgcs.EventColour.GetDiff(x.Value, basePalette.RgbValue) }).ToList();
                var minDistance = colourDistance.Min(x => x.Diff);
                return colourDistance.Find(x => x.Diff == minDistance).Value.Key;
            } catch (System.Exception ex) {
                log.Warn("Failed to get closest Outlook category for " + basePalette.ToString());
                OGCSexception.Analyse(ex);
                return Outlook.OlCategoryColor.olCategoryColorNone;
            }
        }
    }
}
