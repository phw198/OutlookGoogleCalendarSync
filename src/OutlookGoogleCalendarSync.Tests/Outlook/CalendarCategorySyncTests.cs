using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using OutlookCOM = Microsoft.Office.Interop.Outlook;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Outlook {
    [TestClass]
    public class CalendarCategorySyncTests {
        private readonly Ogcs.Outlook.Categories categories = new();

        [TestMethod]
        public void SingleCategoryOnly_SetsOverrideCategoryForCreatedAndUpdatedItems() {
            string createdCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", null, true, "Team calendar", false);
            string updatedCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", "Personal", true, "Team calendar", false);

            Assert.AreEqual("Team calendar", createdCategory);
            Assert.AreEqual("Team calendar", updatedCategory);
        }

        [TestMethod]
        public void CreatedItemsOnly_PreservesExistingOutlookCategory() {
            string createdCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", null, true, "Team calendar", true);
            string updatedCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", "Personal", true, "Team calendar", true);

            Assert.AreEqual("Team calendar", createdCategory);
            Assert.AreEqual("Personal", updatedCategory);
        }

        [TestMethod]
        public void ColourMapping_MultipleCategoriesAllowed_PreservesNonOgcsCategories() {
            string mappedCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", "Personal", false, null, false);
            List<string> updatedCategories = Ogcs.Outlook.Calendar.updateCategoriesForGoogleColour(
                new List<string> { "Personal", "OGCS Red", "Work" }, mappedCategory, false);

            Assert.AreEqual("OGCS Blue", mappedCategory);
            CollectionAssert.AreEqual(new List<string> { "OGCS Blue", "Personal", "Work" }, updatedCategories);
        }

        [TestMethod]
        public void NoCategoryAssignedMapping_MultipleCategoriesAllowed_RemovesOnlyOgcsCategories() {
            string mappedCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory(categories.NO_CATEGORY_ASSIGNED.Key,
                "Personal", false, null, false);
            List<string> updatedCategories = Ogcs.Outlook.Calendar.updateCategoriesForGoogleColour(
                new List<string> { "Personal", "OGCS Blue", "Work", "OGCS Red" }, mappedCategory, false);

            Assert.AreEqual(categories.NO_CATEGORY_ASSIGNED.Key, mappedCategory);
            CollectionAssert.AreEqual(new List<string> { "Personal", "Work" }, updatedCategories);
        }

        [TestMethod]
        public void NoCategoryAssignedMapping_SingleCategoryOnly_RemovesAllCategories() {
            string mappedCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory(categories.NO_CATEGORY_ASSIGNED.Key,
                "Personal", false, null, false);
            List<string> updatedCategories = Ogcs.Outlook.Calendar.updateCategoriesForGoogleColour(
                new List<string> { "Personal", "OGCS Blue", "Work" }, mappedCategory, true);

            Assert.AreEqual(categories.NO_CATEGORY_ASSIGNED.Key, mappedCategory);
            CollectionAssert.AreEqual(new List<string>(), updatedCategories);
        }

        [TestMethod]
        public void NoCategoryAssignedOverride_SingleCategoryOnly_RemovesAllCategories() {
            string overrideCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", "Personal", true,
                Ogcs.Outlook.Calendar.getOverrideCategory(OutlookCOM.OlCategoryColor.olCategoryColorNone, categories.NO_CATEGORY_ASSIGNED.Key), false);
            List<string> updatedCategories = Ogcs.Outlook.Calendar.updateCategoriesForGoogleColour(
                new List<string> { "Personal", "OGCS Blue", "Work" }, overrideCategory, true);

            Assert.AreEqual(categories.NO_CATEGORY_ASSIGNED.Key, overrideCategory);
            CollectionAssert.AreEqual(new List<string>(), updatedCategories);
        }

        [TestMethod]
        public void NoCategoryAssignedOverride_MultipleCategoriesAllowed_RemovesOnlyOgcsCategories() {
            string overrideCategory = Ogcs.Outlook.Calendar.getGoogleToOutlookCategory("OGCS Blue", "Personal", true,
                Ogcs.Outlook.Calendar.getOverrideCategory(OutlookCOM.OlCategoryColor.olCategoryColorNone, categories.NO_CATEGORY_ASSIGNED.Key), false);
            List<string> updatedCategories = Ogcs.Outlook.Calendar.updateCategoriesForGoogleColour(
                new List<string> { "Personal", "OGCS Blue", "Work", "OGCS Red" }, overrideCategory, false);

            Assert.AreEqual(categories.NO_CATEGORY_ASSIGNED.Key, overrideCategory);
            CollectionAssert.AreEqual(new List<string> { "Personal", "Work" }, updatedCategories);
        }

        [TestMethod]
        public void ColourlessOverrideCategory_RetainsConfiguredCategoryName() {
            string overrideCategory = Ogcs.Outlook.Calendar.getOverrideCategory(OutlookCOM.OlCategoryColor.olCategoryColorNone, "Personal");

            Assert.AreEqual("Personal", overrideCategory);
        }
    }
}