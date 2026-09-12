using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Outlook {
    [TestClass]
    public class CategoryFilteringTests {
        private readonly Ogcs.Outlook.Categories categories = new();

        [TestMethod]
        public void ExcludeNoCategoryAssigned_ExcludesItemsWithoutCategories() {
            bool filtered = categories.isExcludedByFilter(null,
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Exclude);

            Assert.IsTrue(filtered);
        }

        [TestMethod]
        public void ExcludeNoCategoryAssigned_IncludesItemsWithOneOrMoreCategories() {
            Assert.IsFalse(categories.isExcludedByFilter("Work",
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Exclude));
            Assert.IsFalse(categories.isExcludedByFilter("Work, Personal",
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Exclude));
        }

        [TestMethod]
        public void IncludeNoCategoryAssigned_IncludesItemsWithoutCategories() {
            bool filtered = categories.isExcludedByFilter(null,
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Include);

            Assert.IsFalse(filtered);
        }

        [TestMethod]
        public void IncludeNoCategoryAssigned_ExcludesItemsWithOneOrMoreCategories() {
            Assert.IsTrue(categories.isExcludedByFilter("Work",
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Include));
            Assert.IsTrue(categories.isExcludedByFilter("Work, Personal",
                new List<string> { categories.NO_CATEGORY_ASSIGNED.Key }, Ogcs.SettingsStore.Calendar.RestrictBy.Include));
        }
    }
}