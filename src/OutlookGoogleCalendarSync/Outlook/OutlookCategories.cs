///#define DEVELOP_AGAINST_2007     //Develop as for Outlook 2007 for greatest compatiblity
using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OutlookCOM = Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Outlook {
    public class Categories {
        public class ColourInfo {
            public String Text { get; }
            public OutlookCOM.OlCategoryColor OutlookCategory { get; }
            public Color Colour { get; }

            public ColourInfo(OutlookCOM.OlCategoryColor category, Color colour, String name = "") {
                this.Text = string.IsNullOrEmpty(name) ? Categories.FriendlyCategoryName(category) : name;
                this.Colour = colour;
                this.OutlookCategory = category;
            }
        }

        public class Map {
            private static readonly ILog log = LogManager.GetLogger(typeof(Map));

            //Source: https://msdn.microsoft.com/en-us/library/ee203806%28v=exchg.80%29.aspx
            public static Dictionary<OutlookCOM.OlCategoryColor, Color> Colours { get; }
            static Map() {
                Colours = new Dictionary<OutlookCOM.OlCategoryColor, Color> {
                { OutlookCOM.OlCategoryColor.olCategoryColorBlack, Color.FromArgb(28,28,28) },
                { OutlookCOM.OlCategoryColor.olCategoryColorBlue, Color.FromArgb(50, 103, 184) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkBlue, Color.FromArgb(42, 81, 145 ) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkGray, Color.FromArgb(165, 165, 165) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkGreen, Color.FromArgb(53, 121, 43) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkMaroon, Color.FromArgb(130, 55, 95) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkOlive, Color.FromArgb(95, 108, 58) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkOrange, Color.FromArgb(177, 79, 13 ) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkPeach, Color.FromArgb(171, 123, 5 ) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkPurple, Color.FromArgb(80, 50, 143) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkRed, Color.FromArgb(175, 30, 37) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkSteel, Color.FromArgb(140, 156, 189) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkTeal, Color.FromArgb(46, 125, 100) },
                { OutlookCOM.OlCategoryColor.olCategoryColorDarkYellow, Color.FromArgb(153, 148, 0) },
                { OutlookCOM.OlCategoryColor.olCategoryColorGray, Color.FromArgb(196, 196, 196) },
                { OutlookCOM.OlCategoryColor.olCategoryColorGreen, Color.FromArgb(74, 182, 63) },
                { OutlookCOM.OlCategoryColor.olCategoryColorMaroon, Color.FromArgb(163, 78, 120) },
                { OutlookCOM.OlCategoryColor.olCategoryColorNone, Color.FromArgb(255, 255, 255) },
                { OutlookCOM.OlCategoryColor.olCategoryColorOlive, Color.FromArgb(133, 154, 82) },
                { OutlookCOM.OlCategoryColor.olCategoryColorOrange, Color.FromArgb(240, 108, 21) },
                { OutlookCOM.OlCategoryColor.olCategoryColorPeach, Color.FromArgb(255, 202, 76) },
                { OutlookCOM.OlCategoryColor.olCategoryColorPurple, Color.FromArgb(97, 61, 180) },
                { OutlookCOM.OlCategoryColor.olCategoryColorRed, Color.FromArgb(214, 37, 46) },
                { OutlookCOM.OlCategoryColor.olCategoryColorSteel, Color.FromArgb(196, 204, 221) },
                { OutlookCOM.OlCategoryColor.olCategoryColorTeal, Color.FromArgb(64, 189, 149) },
                { OutlookCOM.OlCategoryColor.olCategoryColorYellow, Color.FromArgb(255, 254, 61) }
            };
            }

            /// <summary>
            /// Convert from Outlook category colour to Color
            /// </summary>
            public static Color RgbColour(OutlookCOM.OlCategoryColor colour) {
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

            public static OutlookCOM.OlCategoryColor GetClosestCategory(Ogcs.Google.EventColour.Palette basePalette) {
                try {
                    var colourDistance = Colours.Select(x => new { Value = x, Diff = Ogcs.Google.EventColour.GetDiff(x.Value, basePalette.RgbValue) }).ToList();
                    var minDistance = colourDistance.Min(x => x.Diff);
                    return colourDistance.Find(x => x.Diff == minDistance).Value.Key;
                } catch (System.Exception ex) {
                    log.Warn("Failed to get closest Outlook category for " + basePalette.ToString());
                    OGCSexception.Analyse(ex);
                    return OutlookCOM.OlCategoryColor.olCategoryColorNone;
                }
            }
        }

        private static readonly ILog log = LogManager.GetLogger(typeof(Categories));

        private OutlookCOM.Categories _categories;
        private OutlookCOM.Categories categories {
            get {
                try {
                    int? testAccess = this._categories?.Count;
                } catch {
                    this._categories = null;
                }
                if (_categories == null) ValidateCategories();
                return _categories;
            }
            set { _categories = value; }
        }

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
            _categories = (OutlookCOM.Categories)Calendar.ReleaseObject(_categories);
        }

        /// <summary>
        /// Get the categories as configured in Outlook and store in class
        /// </summary>
        /// <param name="oApp"></param>
        /// <param name="calendar"></param>
        public void Get(OutlookCOM.Application oApp, OutlookCOM.MAPIFolder calendar) {
            OutlookCOM.Store store = null;
            try {
                if (Settings.Profile.InPlay().OutlookService == Calendar.Service.DefaultMailbox)
                    this.categories = oApp.Session.Categories;
                else {
                    try {
                        store = calendar.Store;
                        if (Factory.OutlookVersionName == Factory.OutlookVersionNames.Outlook2007) {
                            log.Debug("Accessing Outlook 2007 categories via reflection.");
                            this.categories = store.GetType().GetProperty("Categories").GetValue(store, null) as OutlookCOM.Categories;
                        } else {
#if !DEVELOP_AGAINST_2007
                            log.Debug("Accessing categories through Outlook 2010 store.");
                            this.categories = store.Categories;
#else
                            log.Debug("Accessing Outlook 2007 categories via reflection.");
                            this.categories = store.GetType().GetProperty("Categories").GetValue(store, null) as OutlookCOM.Categories;
#endif
                        }
                    } catch (System.Exception ex) {
                        Outlook.Errors.ErrorType error = Outlook.Errors.HandleComError(ex);
                        if (error == Outlook.Errors.ErrorType.RpcServerUnavailable || error == Outlook.Errors.ErrorType.WrongThread) 
                            throw;

                        log.Warn("Failed getting non-default mailbox categories. " + ex.Message);
                        log.Debug("Reverting to default mailbox categories.");
                        this.categories = oApp.Session.Categories;
                    }
                }
            } finally {
                store = (OutlookCOM.Store)Calendar.ReleaseObject(store);
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
        public OutlookCOM.OlCategoryColor? OutlookColour(String categoryName) {
            if (string.IsNullOrEmpty(categoryName)) log.Warn("Category name is empty.");

            foreach (OutlookCOM.Category category in this.categories) {
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
                if (this._categories != null && this._categories.Count > 0) { }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Categories are not accessible!", OGCSexception.LogAsFail(ex));
                this._categories = null;
            }
            if (this._categories == null || Calendar.Categories == null) {
                Calendar.Instance.IOutlook.RefreshCategories();
            } else {
                String catName = "";
                try {
                    foreach (OutlookCOM.Category cat in this._categories) {
                        catName = cat.Name;
                        OutlookCOM.OlCategoryColor catColour = cat.Color;
                    }
                } catch (System.Exception ex) {
                    if (OGCSexception.GetErrorCode(ex, 0x0000FFFF) == "0x00004005" && ex.TargetSite.Name == "get_Color") {
                        log.Warn("Outlook category '" + catName + "' seems to have changed!");
                        Calendar.Instance.IOutlook.RefreshCategories();
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
            if (this._categories != null) {
                foreach (OutlookCOM.Category category in this.categories) {
                    names.Add(category.Name);
                }
            }
            return names;
        }

        /// <summary>
        /// Get the Outlook categories as List of ColourInfo
        /// </summary>
        /// <returns>List to be used in dropdown, for example</returns>
        public List<Categories.ColourInfo> DropdownItems() {
            List<Categories.ColourInfo> items = new List<Categories.ColourInfo>();
            foreach (OutlookCOM.Category category in this.categories) {
                items.Add(new Categories.ColourInfo(category.Color, Categories.Map.RgbColour(category.Color), category.Name));
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
        public String FindName(OutlookCOM.OlCategoryColor? olCategory, String categoryName = null, Boolean createMissingCategory = true) {
            if (olCategory == null || olCategory == OutlookCOM.OlCategoryColor.olCategoryColorNone) return "";

            OutlookCOM.Category failSafeCategory = null;
            foreach (OutlookCOM.Category category in this.categories) {
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
                        Calendar.Instance.IOutlook.RefreshCategories();
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
                OutlookCOM.Category newCategory = categories.Add(newCategoryName, olCategory);
                log.Info("Added new Outlook category \"" + newCategory.Name + "\" for " + newCategory.Color.ToString());
                return newCategory.Name;
            }
            return "";
        }

        public static String FriendlyCategoryName(OutlookCOM.OlCategoryColor? olCategory) {
            return olCategory.ToString().Replace("olCategoryColor", "").Replace("Dark", "Dark ");
        }
    }
}
