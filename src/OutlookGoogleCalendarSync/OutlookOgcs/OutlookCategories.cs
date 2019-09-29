using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public class Category {
        public String Name { get; internal set; }
        public Outlook.OlCategoryColor Colour { get; internal set; }

        public Category(String name, Outlook.OlCategoryColor colour) {
            this.Name = name;
            this.Colour = colour;
        }
    }

    public class Categories {
        private static readonly ILog log = LogManager.GetLogger(typeof(Categories));
        private Outlook.Categories oomCategories;
        private List<OutlookOgcs.Category> categories;
        public String Delimiter { get; }
        private const string categoryListPropertySchemaName = @"http://schemas.microsoft.com/mapi/proptag/0x7C080102";
        //private Boolean usingStorageXmlCategories = false;

        public Categories() {
            try {
                oomCategories = null;
                categories = new List<OutlookOgcs.Category>();
                Delimiter = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator + " ";
                //usingStorageXmlCategories = false;
            } catch (System.Exception ex) {
                log.Error("Failed to get system ListSeparator value.");
                OGCSexception.Analyse(ex);
                Delimiter = ", ";
            }
        }

        public void Dispose() {
            oomCategories = (Outlook.Categories)OutlookOgcs.Calendar.ReleaseObject(oomCategories);
        }

        private void loadDefaultCategories(Outlook.Categories outlookCategories) {
            log.Debug("Loading default categories from Outlook.");
            oomCategories = outlookCategories;
            foreach (Outlook.Category cat in outlookCategories) {
                this.categories.Add(new Category(cat.Name, cat.Color));
            }
        }

        /// <summary>
        /// Get the categories as configured in Outlook and store in class
        /// </summary>
        /// <param name="oApp"></param>
        /// <param name="calendar"></param>
        public void Get(Outlook.Application oApp, Outlook.MAPIFolder calendar) {
            categories.Clear();
            if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.DefaultMailbox) {
                GetFromStorage(calendar);
                this.loadDefaultCategories(oApp.Session.Categories);                
            } else {
                Outlook.Store store = null;
                try {
                    log.Debug("Retrieving store.");
                    store = calendar.Store;
                    System.Reflection.PropertyInfo pi = store.GetType().GetProperty("Categories");
                    if (pi == null) {
                        log.Warn("'Categories' property is null! Listing all properties...");
                        System.Reflection.PropertyInfo[] pis = store.GetType().GetProperties();
                        if (pis == null || pis.Count() == 0) log.Warn("There are no properties at all!");
                        else pis.ToList().ForEach(p => log.Debug(p.Name));
                    }
                    this.loadDefaultCategories(store.GetType().GetProperty("Categories").GetValue(store, null) as Outlook.Categories);                        

                } catch (System.Exception ex) {
                    log.Warn("Failed getting non-default mailbox categories by reflection. " + ex.Message);
                    oomCategories = (Outlook.Categories)OutlookOgcs.Calendar.ReleaseObject(oomCategories);
                    if (!GetFromStorage(calendar)) {
                        log.Debug("Reverting to default mailbox categories.");
                        oomCategories = oApp.Session.Categories;
                        this.loadDefaultCategories(oomCategories);
                    }
                } finally {
                    store = (Outlook.Store)Calendar.ReleaseObject(store);
                }
            }
        }

        private Boolean GetFromStorage(Outlook.MAPIFolder calendar) {
            Outlook.StorageItem categoryStorage = null;

            try {
                categoryStorage = calendar.GetStorage("IPM.Configuration.CategoryList", Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);

                if (categoryStorage != null) {
                    Outlook.PropertyAccessor categoryPA = null;
                    System.Xml.XmlReader xmlReader = null;
                    try {
                        categoryPA = categoryStorage.PropertyAccessor;
                        var xmlBytes = (byte[])categoryPA.GetProperty(categoryListPropertySchemaName);
                        String xmlString = System.Text.Encoding.UTF8.GetString(xmlBytes);
                        log.Debug(xmlString);
                        xmlReader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlString));
                        xmlReader.ReadToFollowing("category");

                        while (!xmlReader.EOF) {
                            try {
                                if (xmlReader.NodeType == System.Xml.XmlNodeType.Whitespace) continue;
                                if (xmlReader.NodeType == System.Xml.XmlNodeType.Element && xmlReader.Name == "category") {
                                    String name = xmlReader.GetAttribute("name").ToString();
                                    String colour = xmlReader.GetAttribute("color").ToString();
                                    Outlook.OlCategoryColor category = (Outlook.OlCategoryColor)(Convert.ToInt16(colour) + 1);
                                    this.categories.Add(new Category(name, category));
                                }
                            } finally {
                                xmlReader.Read();
                            }
                        }
                        return true;

                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Could not read categories from storage object.", ex);
                    } finally {
                        if (xmlReader != null) xmlReader.Close();
                        categoryPA = (Outlook.PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(categoryPA);
                    }
                } else
                    log.Warn("No categories found in storage.");

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not open category list from storage.", ex);
            } finally {
                categoryStorage = (Outlook.StorageItem)OutlookOgcs.Calendar.ReleaseObject(categoryStorage);
            }
            return false;
        }

        public void BuildPicker(ref System.Windows.Forms.CheckedListBox clb) {
            clb.BeginUpdate();
            clb.Items.Clear();
            clb.Items.Add("<No category assigned>");
            foreach (String catName in getNames()) {
                clb.Items.Add(catName);
            }
            foreach (String cat in Settings.Instance.Categories) {
                try {
                    clb.SetItemChecked(clb.Items.IndexOf(cat), true);
                } catch { /* Category "cat" no longer exists */ }
            }
            clb.EndUpdate();
        }

        /// <summary>
        /// Get the Outlook category from the name given to the category
        /// </summary>
        /// <param name="categoryName">The user named Outlook category</param>
        /// <returns>The Outlook category type</returns>
        public Outlook.OlCategoryColor? OutlookColour(String categoryName) {
            if (string.IsNullOrEmpty(categoryName)) log.Warn("Category name is empty.");

            foreach (OutlookOgcs.Category category in this.categories) {
                if (category.Name == categoryName.Trim()) return category.Colour;
            }
            return null;
        }

        /// <summary>
        /// Return all the category names as a list of strings.
        /// </summary>
        private List<String> getNames() {
            List<String> names = new List<String>();
            if (this.categories != null) {
                foreach (OutlookOgcs.Category category in this.categories) {
                    names.Add(category.Name);
                }
            }
            return names;
        }

        /// <summary>
        /// Get the Outlook categories as List of ColourInfo
        /// </summary>
        /// <returns>List to be used in dropdown, for example</returns>
        public List<Extensions.ColourPicker.ColourInfo> DropdownItems() {
            List<Extensions.ColourPicker.ColourInfo> items = new List<Extensions.ColourPicker.ColourInfo>();
            if (this.categories != null) {
                foreach (OutlookOgcs.Category category in this.categories) {
                    items.Add(new Extensions.ColourPicker.ColourInfo(category.Colour, CategoryMap.RgbColour(category.Colour), category.Name));
                }
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

            OutlookOgcs.Category failSafeCategory = null;
            foreach (OutlookOgcs.Category category in this.categories) {
                if (category.Colour == olCategory) {
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

            String newCategoryName = "OGCS " + FriendlyCategoryName(olCategory);
            if (oomCategories != null) {
                Outlook.Category newCategory = oomCategories.Add(newCategoryName, olCategory);
            } else {
                addNewStorageCategory(null, newCategoryName, olCategory);
            }
            categories.Add(new OutlookOgcs.Category(newCategoryName, olCategory));
            log.Info("Added new Outlook category \"" + newCategoryName + "\" for " + olCategory.ToString());
            return newCategoryName;
        }

        public static String FriendlyCategoryName(Outlook.OlCategoryColor olCategory) {
            return olCategory.ToString().Replace("olCategoryColor", "").Replace("Dark", "Dark ");
        }

        private void addNewStorageCategory(Outlook.MAPIFolder calendar, String categoryName, Outlook.OlCategoryColor categoryColour) {
            Outlook.StorageItem categoryStorage = null;

            try {
                categoryStorage = calendar.GetStorage("IPM.Configuration.CategoryList", Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);
                if (categoryStorage != null) {
                    Outlook.PropertyAccessor categoryPA = null;
                    try {
                        categoryPA = categoryStorage.PropertyAccessor;
                        byte[] xmlBytes = (byte[])categoryPA.GetProperty(categoryListPropertySchemaName);
                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.LoadXml(System.Text.Encoding.UTF8.GetString(xmlBytes));

                        System.Xml.XmlNodeList nl = xmlDoc.GetElementsByTagName("category");
                        System.Xml.XmlNode lastNode = nl[nl.Count - 1];
                        System.Xml.XmlNode newNode = xmlDoc.CreateElement("category");
                        newNode = lastNode.Clone();
                        newNode.Attributes["name"].Value = categoryName;
                        newNode.Attributes["color"].Value = ((int)categoryColour - 1).ToString();
                        newNode.Attributes["keyboardShortcut"].Value = "0";

                        xmlDoc.DocumentElement.AppendChild(newNode);
                        xmlBytes = System.Text.Encoding.UTF8.GetBytes(xmlDoc.InnerXml);
                        categoryPA.SetProperty(categoryListPropertySchemaName, xmlBytes);
                        categoryStorage.Save();
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Not able to add new category to storage XML.", ex);
                    } finally {
                        categoryPA = (Outlook.PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(categoryPA);
                    }
                } else
                    log.Warn("No categories found in storage.");

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not open category list from storage.", ex);
            } finally {
                categoryStorage = (Outlook.StorageItem)OutlookOgcs.Calendar.ReleaseObject(categoryStorage);
            }
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
