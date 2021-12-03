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
                OGCSexception.Analyse(ex);
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
                if (OGCSexception.GetErrorCode(ex) == "0x80131509") { //Sequence contains no elements
                    log.Warn("'" + nodeName + "' could not be found.");
                } else {
                    log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                    OGCSexception.Analyse(ex);
                }
                return null;

            } catch (System.IO.IOException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80070020") { //Setting file in use by another process
                    log.Warn("Failed retrieving '" + nodeName + "' from " + filename);
                    log.Warn(ex.Message);
                } else {
                    log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                    OGCSexception.Analyse(ex);
                }
                return null;

            } catch (System.Xml.XmlException ex) {
                log.Warn("Failed retrieving '" + nodeName + "' from " + filename);
                if (OGCSexception.GetErrorCode(ex) == "0x80131940") { //hexadecimal value 0x00, is an invalid character
                    log.Warn(ex.Message);
                    Settings.ResetFile(filename);
                    try {
                        XDocument xml = XDocument.Load(filename);
                        XElement settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
                        XElement xe = settingsXE.Elements(ns + nodeName).First();
                        if (debugLog)
                            log.Debug("Retrieved setting '" + nodeName + "' with value '" + xe.Value + "'");
                        return xe.Value;
                    } catch (System.Exception ex2) {
                        OGCSexception.Analyse("Still failed to retrieve '" + nodeName + "' from " + filename, ex2);
                        return null;
                    }
                } else {
                    OGCSexception.Analyse(ex);
                    return null;
                }

            } catch (System.Exception ex) {
                log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                OGCSexception.Analyse(ex);
                return null;
            }
        }
        
        public static void ExportElement(string nodeName, object nodeValue, string filename) {
            XDocument xml = null;
            try {
                xml = XDocument.Load(filename);
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to load " + filename, ex, true);
                throw;
            }
            XElement settingsXE = null;
            try {
                settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
            } catch (System.Exception ex) {
                log.Debug(filename + " head: " + xml.ToString().Substring(0, Math.Min(200, xml.ToString().Length)));
                OGCSexception.Analyse("Could not access 'Settings' element.", ex, true);
                return;
            }
            try {
                XElement xe = settingsXE.Elements(ns + nodeName).First();
                if (nodeValue == null && nodeName == "CloudLogging") { //Nullable Boolean node(s)
                    XNamespace i = "http://www.w3.org/2001/XMLSchema-instance";
                    xe.SetAttributeValue(i + "nil", "true"); //Add nullable attribute 'i:nil="true"'
                    xe.SetValue(String.Empty);
                }  else {
                    xe.SetValue(nodeValue);
                    if (nodeValue is Boolean && nodeValue != null)
                        xe.RemoveAttributes(); //Remove nullable attribute 'i:nil="true"'
                }
                xml.Save(filename);
                log.Debug("Setting '" + nodeName + "' updated to '" + nodeValue + "'");
            } catch (System.Exception ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80131509") { //Sequence contains no elements
                    log.Debug("Adding Setting " + nodeName + " to settings.xml");
                    settingsXE.Add(new XElement(ns + nodeName, nodeValue));
                    xml.Root.Sort();
                    xml.Save(filename);
                } else {
                    OGCSexception.Analyse("Failed to export setting " + nodeName + "=" + nodeValue + " to " + filename + " file.", ex);
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
            List<XElement> sortedChildren = source.Elements().OrderBy(e => e.Name.ToString()).ToList();
            if (source.HasElements) {
                source.RemoveNodes();
                sortedChildren.ForEach(c => c.Sort(bSortAttributes));
                sortedChildren.ForEach(c => source.Add(c));
            }
        }
    }
}
