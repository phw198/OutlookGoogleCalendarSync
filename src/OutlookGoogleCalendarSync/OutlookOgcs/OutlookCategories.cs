using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public class Categories {
        public class ColourInfo {
            public String Text { get; }
            public Outlook.OlCategoryColor OutlookCategory { get; }
            public Color Colour { get; }

            public ColourInfo(Outlook.OlCategoryColor category, Color colour, String name = "") {
                this.Text = string.IsNullOrEmpty(name) ? OutlookOgcs.Categories.FriendlyCategoryName(category) : name;
                this.Colour = colour;
                this.OutlookCategory = category;
            }
        }

        public class Map {
            private static readonly ILog log = LogManager.GetLogger(typeof(Map));

            //Source: https://msdn.microsoft.com/en-us/library/ee203806%28v=exchg.80%29.aspx
            public static Dictionary<Outlook.OlCategoryColor, Color> Colours { get; }
            static Map() {
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
                Color colour = new Color();
                try {
                    colour = ColorTranslator.FromHtml(hexColour);
                    log.Fine("Converted '" + hexColour + "' to " + colour.ToString());
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Could not convert hex '" + hexColour + "' to RGB colour.", ex);
                }
                return colour;
            }

            public static Outlook.OlCategoryColor GetClosestCategory(GoogleOgcs.EventColour.Palette basePalette) {
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

        public void Dispose() {
            categories = (Outlook.Categories)OutlookOgcs.Calendar.ReleaseObject(categories);
        }

        /// <summary>
        /// Get the categories as configured in Outlook and store in class
        /// </summary>
        /// <param name="oApp"></param>
        /// <param name="calendar"></param>
        public void Get(Outlook.Application oApp, Outlook.MAPIFolder calendar) {
            Outlook.Store store = null;
            try {
                if (Settings.Profile.InPlay().OutlookService == OutlookOgcs.Calendar.Service.DefaultMailbox)
                    this.categories = oApp.Session.Categories;
                else {
                    try {
                        store = calendar.Store;
                        this.categories = store.GetType().GetProperty("Categories").GetValue(store, null) as Outlook.Categories;
                    } catch (System.Exception ex) {
                        log.Warn("Failed getting non-default mailbox categories. " + ex.Message);
                        log.Debug("Reverting to default mailbox categories.");
                        this.categories = oApp.Session.Categories;
                    }
                }
            } finally {
                store = (Outlook.Store)Calendar.ReleaseObject(store);
            }
        }

        public void BuildPicker(ref System.Windows.Forms.CheckedListBox clb) {
            clb.BeginUpdate();
            clb.Items.Clear();
            clb.Items.Add("<No category assigned>");
            foreach (String catName in getNames()) {
                clb.Items.Add(catName);
            }
            foreach (String cat in Forms.Main.Instance.ActiveCalendarProfile.Categories) {
                try {
                    clb.SetItemChecked(clb.Items.IndexOf(cat), true);
                } catch { /* Category "cat" no longer exists */ }
            }
            clb.EndUpdate();
        }

        /// <summary>
        /// Get the Outlook category colour from the name given to the category
        /// </summary>
        /// <param name="categoryName">The user named Outlook category</param>
        /// <returns>The Outlook category type</returns>
        public Outlook.OlCategoryColor? OutlookColour(String categoryName) {
            if (string.IsNullOrEmpty(categoryName)) log.Warn("Category name is empty.");

            foreach (Outlook.Category category in this.categories) {
                if (category.Name == categoryName.Trim()) return category.Color;
            }

            log.Warn("Could not convert category name '" + categoryName + "' into Outlook category type.");
            return null;
        }

        /// <summary>
        /// Check all the Outlook categories can still be accessed. If not, refresh them.
        /// </summary>
        public void ValidateCategories() {
            try {
                if (this.categories != null && this.categories.Count > 0) { }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Categories are not accessible!", OGCSexception.LogAsFail(ex));
                this.categories = null;
            }
            if (this.categories == null || OutlookOgcs.Calendar.Categories == null) {
                OutlookOgcs.Calendar.Instance.IOutlook.RefreshCategories();
            } else {
                String catName = "";
                try {
                    foreach (Outlook.Category cat in this.categories) {
                        catName = cat.Name;
                        Outlook.OlCategoryColor catColour = cat.Color;
                    }
                } catch (System.Exception ex) {
                    if (OGCSexception.GetErrorCode(ex, 0x0000FFFF) == "0x00004005" && ex.TargetSite.Name == "get_Color") {
                        log.Warn("Outlook category '" + catName + "' seems to have changed!");
                        OutlookOgcs.Calendar.Instance.IOutlook.RefreshCategories();
                    } else {
                        OGCSexception.Analyse("Could not access all the Outlook categories.", ex);
                        this.categories = null;
                        return;
                    }
                }
            }
            this.categories = Calendar.Categories.categories;
        }

        /// <summary>
        /// Return all the category names as a list of strings.
        /// </summary>
        private List<String> getNames() {
            List<String> names = new List<String>();
            if (this.categories != null) {
                foreach (Outlook.Category category in this.categories) {
                    names.Add(category.Name);
                }
            }
            return names;
        }

        /// <summary>
        /// Get the Outlook categories as List of ColourInfo
        /// </summary>
        /// <returns>List to be used in dropdown, for example</returns>
        public List<OutlookOgcs.Categories.ColourInfo> DropdownItems() {
            List<OutlookOgcs.Categories.ColourInfo> items = new List<OutlookOgcs.Categories.ColourInfo>();
            if (this.categories != null) {
                foreach (Outlook.Category category in this.categories) {
                    items.Add(new OutlookOgcs.Categories.ColourInfo(category.Color, Categories.Map.RgbColour(category.Color), category.Name));
                }
            }
            return items.OrderBy(i => i.Text).ToList();
        }

        /// <summary>
        /// Get the Outlook category name for a given category colour.
        /// If category not yet used, new one added of the form "OGCS [colour]"
        /// </summary>
        /// <param name="olCategory">The Outlook category to search by</param>
        /// <param name="categoryName">Optional: The Outlook category name to also search by</param>
        /// <param name="createMissingCategory">Optional: Create unused category colour?</param>
        /// <returns>The matching category name</returns>
        public String FindName(Outlook.OlCategoryColor? olCategory, String categoryName = null, Boolean createMissingCategory = true) {
            if (olCategory == null || olCategory == Outlook.OlCategoryColor.olCategoryColorNone) return "";

            Outlook.Category failSafeCategory = null;
            foreach (Outlook.Category category in this.categories) {
                try {
                    if (category.Color == olCategory) {
                        if (categoryName == null) {
                            if (category.Name.StartsWith("OGCS ")) return category.Name;
                            else if (!createMissingCategory) return category.Name;
                        } else {
                            if (category.Name == categoryName) return category.Name;
                            if (category.Name.StartsWith("OGCS ")) failSafeCategory = category;
                        }
                    }
                } catch (System.Runtime.InteropServices.COMException ex) {
                    if (OGCSexception.GetErrorCode(ex, 0x0000FFFF) == "0x00004005") { //The operation failed.
                        log.Warn("It seems a category has been manually removed in Outlook.");
                        OutlookOgcs.Calendar.Instance.IOutlook.RefreshCategories();
                    } else throw;
                }
            }

            if (failSafeCategory != null) {
                log.Warn("Failed to find Outlook category " + olCategory.ToString() + " with name '" + categoryName + "'");
                log.Debug("Using category with name \"" + failSafeCategory.Name + "\" instead.");
                return failSafeCategory.Name;
            }

            log.Debug("Did not find Outlook category " + olCategory.ToString() + (categoryName == null ? "" : " \"" + categoryName + "\""));
            String newCategoryName = "OGCS " + FriendlyCategoryName(olCategory);
            if (!createMissingCategory) {
                createMissingCategory = OgcsMessageBox.Show("There is no matching Outlook category.\r\nWould you like to create one of the form '" + newCategoryName + "'?",
                    "Create new Outlook category?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            }
            if (createMissingCategory) {
                Outlook.Category newCategory = categories.Add(newCategoryName, olCategory);
                log.Info("Added new Outlook category \"" + newCategory.Name + "\" for " + newCategory.Color.ToString());
                return newCategory.Name;
            }
            return "";
        }

        public static String FriendlyCategoryName(Outlook.OlCategoryColor? olCategory) {
            return olCategory.ToString().Replace("olCategoryColor", "").Replace("Dark", "Dark ");
        }
    }
}
