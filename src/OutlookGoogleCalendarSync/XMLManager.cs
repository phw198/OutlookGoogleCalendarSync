using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Linq;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Exports or imports any object to/from XML.
    /// </summary>
    public static class XMLManager {
        private static readonly ILog log = LogManager.GetLogger(typeof(XMLManager));
        private static XNamespace ns = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync";
        
        /// <summary>
        /// Exports any object given in "obj" to an xml file given in "filename"
        /// </summary>
        /// <param name="obj">The object that is to be serialized/exported to XML.</param>
        /// <param name="filename">The filename of the xml file to be written.</param>
        public static void Export(Object obj, string filename) {
            XmlTextWriter writer = new XmlTextWriter(filename, null) {
                Formatting = Formatting.Indented,
                Indentation = 4
            };
            new DataContractSerializer(obj.GetType()).WriteObject(writer, obj);
            writer.Close();
        }
        
        /// <summary>
        /// Imports from XML and returns the resulting object of type T.
        /// </summary>
        /// <param name="filename">The XML file from which to import.</param>
        /// <returns></returns>
        public static T Import<T>(string filename) {
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            T result = default(T);
            try {
                result = (T)new DataContractSerializer(typeof(T)).ReadObject(fs);
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
                if (Forms.Main.Instance != null) Forms.Main.Instance.tabApp.SelectedTab = Forms.Main.Instance.tabPage_Settings;
                throw new ApplicationException("Failed to import settings.");
            } finally {
                fs.Close();
            }
            return result;
        }

        public static string ImportElement(string nodeName, string filename, Boolean debugLog = true) {
            try {
                XDocument xml = XDocument.Load(filename);
                XElement settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
                XElement xe = settingsXE.Elements(ns + nodeName).First();
                if (debugLog)
                    log.Debug("Retrieved setting '" + nodeName + "' with value '" + xe.Value + "'");
                return xe.Value;

            } catch (System.InvalidOperationException ex) {
                if (ex.GetErrorCode() == "0x80131509") { //Sequence contains no elements
                    log.Warn("'" + nodeName + "' could not be found.");
                } else {
                    log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                    Ogcs.Exception.Analyse(ex);
                }
                return null;

            } catch (System.IO.IOException ex) {
                if (ex.GetErrorCode() == "0x80070020") { //Setting file in use by another process
                    log.Warn("Failed retrieving '" + nodeName + "' from " + filename);
                    log.Warn(ex.Message);
                } else {
                    log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                    Ogcs.Exception.Analyse(ex);
                }
                return null;

            } catch (System.Xml.XmlException ex) {
                log.Warn($"Failed retrieving '{nodeName}' from {filename}");
                if (ex.GetErrorCode() == "0x80131940") { //hexadecimal value 0x00, is an invalid character
                    log.Warn(ex.Message);
                    Settings.ResetFile(filename);
                    try {
                        XDocument xml = XDocument.Load(filename);
                        XElement settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
                        XElement xe = settingsXE.Elements(ns + nodeName).First();
                        if (debugLog)
                            log.Debug($"Retrieved setting '{nodeName}' with value '{xe.Value}'");
                        return xe.Value;
                    } catch (System.Exception ex2) {
                        ex2.Analyse($"Still failed to retrieve '{nodeName}' from {filename}");
                        return null;
                    }
                } else {
                    Ogcs.Exception.Analyse(ex);
                    return null;
                }

            } catch (System.Exception ex) {
                log.Error($"Failed retrieving '{nodeName}' from {filename}");
                Ogcs.Exception.Analyse(ex);
                return null;
            }
        }
        
        public static void ExportElement(Object settingStore, String nodeName, object nodeValue, string filename) {
            XDocument xml = null;
            try {
                xml = XDocument.Load(filename);
            } catch (System.Exception ex) {
                ex.Analyse("Failed to load " + filename, true);
                throw;
            }
            XElement settingsXE = null;
            try {
                settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
            } catch (System.Exception ex) {
                log.Debug(filename + " head: " + xml.ToString().Substring(0, Math.Min(200, xml.ToString().Length)));
                ex.Analyse("Could not access 'Settings' element.", true);
                return;
            }
            XElement xe = null;
            XElement xeProfile = null;
            try {
                if (Settings.Profile.GetType(settingStore) == Settings.Profile.Type.Calendar) {
                    //It's a Calendar setting
                    SettingsStore.Calendar calSettings = settingStore as SettingsStore.Calendar;
                    XElement xeCalendars = settingsXE.Elements(ns + "Calendars").First();
                    List<XElement> xeCalendar = xeCalendars.Elements(ns + Settings.Profile.Type.Calendar.ToString()).ToList();
                    xeProfile = xeCalendar.First(c => c.Element(ns + "_ProfileName").Value == calSettings._ProfileName);
                    xe = xeProfile.Elements(ns + nodeName).First();
                } else if (settingStore is Settings) {
                    //It's a "global" setting
                    xe = settingsXE.Elements(ns + nodeName).First();
                }
                if (nodeValue == null && nodeName == "CloudLogging") { //Nullable Boolean node(s)
                    XNamespace i = "http://www.w3.org/2001/XMLSchema-instance";
                    xe.SetAttributeValue(i + "nil", "true"); //Add nullable attribute 'i:nil="true"'
                    xe.SetValue(String.Empty);
                } else {
                    xe.SetValue(nodeValue);
                    if (nodeValue is Boolean && nodeValue != null)
                        xe.RemoveAttributes(); //Remove nullable attribute 'i:nil="true"'
                }
                xml.Save(filename);
                log.Debug("Setting '" + nodeName + "' updated to '" + nodeValue + "'");
            } catch (System.Exception ex) {
                if (ex.GetErrorCode() == "0x80131509") { //Sequence contains no elements
                    log.Debug("Adding Setting " + nodeName + " for " + settingStore.ToString() + " to " + filename);
                    if (xeProfile != null)
                        xeProfile.Add(new XElement(ns + nodeName, nodeValue));
                    else
                        settingsXE.Add(new XElement(ns + nodeName, nodeValue));
                    xml.Root.Sort();
                    xml.Save(filename);
                } else {
                    ex.Analyse($"Failed to export setting {nodeName}={nodeValue} for {settingStore.ToString()} to {filename} file.");
                }
            }
        }

        public static void Sort(this XElement source, bool bSortAttributes = true) {
            //Make sure there is a valid source
            if (source == null) throw new ArgumentNullException("source");

            //Sort attributes if needed
            if (bSortAttributes) {
                List<XAttribute> sortedAttributes = source.Attributes().OrderBy(a => a.ToString()).ToList();
                sortedAttributes.ForEach(a => a.Remove());
                sortedAttributes.ForEach(a => source.Add(a));
            }

            //Sort the children if any exist
            List<XElement> sortedChildren = source.Elements().OrderBy(e => e.Name.ToString(), StringComparer.Ordinal).ToList();
            if (source.HasElements) {
                if (source.Name.LocalName == "ColourMaps") return; //This is a dictionary element, and non-alphabetic order must be maintained.
                source.RemoveNodes();
                sortedChildren.ForEach(c => c.Sort(bSortAttributes));
                sortedChildren.ForEach(c => source.Add(c));
            }
        }
        
        public static XElement GetElement(String needleElement, XDocument xmlDoc) {
            return xmlDoc.Element(ns + needleElement);
        }
        private static XElement getElement(String needleElement, XElement element) {
            return element.Element(ns + needleElement);
        }

        public static XElement AddElement(String nodeName, XElement parent, String value = null, Boolean onlySingleInstance = true) {
            log.Debug("Adding element '" + nodeName + "' under '" + parent.Name.LocalName + "'");
            XElement alreadyExists = getElement(nodeName, parent);
            if (alreadyExists != null ) {
                if (onlySingleInstance) {
                    log.Warn("'" + nodeName + "' already exists. Not adding another.");
                    return alreadyExists;
                }
                log.Debug("'" + nodeName + "' already exists. Adding another.");
            }
            parent.Add(new XElement(ns + nodeName, value));
            return getElement(nodeName, parent);
        }

        private static void removeElement(String nodeName, XElement parent) {
            log.Debug("Removing element '" + nodeName + "' under '" + parent.Name.LocalName + "'");
            XElement target = getElement(nodeName, parent);
            if (target == null)
                log.Warn("Could not find element '" + nodeName + "' under '" + parent.Name.LocalName + "'");
            else
                target.Remove();
        }

        /// <summary>
        /// Fails silently if node to be moved does not exist.
        /// </summary>
        /// <param name="nodeName">Node to be moved</param>
        /// <param name="parent">The parent of node being moved</param>
        /// <param name="target">New parent</param>
        public static void MoveElement(String nodeName, XElement parent, XElement target) {
            try {
                XElement sourceElement = getElement(nodeName, parent);
                target.Add(sourceElement);
                removeElement(nodeName, parent);
            } catch (System.Exception ex) {
                ex.Analyse($"Could not move '{nodeName}'");
            }
        }
    }
}
