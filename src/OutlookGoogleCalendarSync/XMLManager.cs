using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Linq;
using log4net;

namespace OutlookGoogleCalendarSync
{
    /// <summary>
    /// Exports or imports any object to/from XML.
    /// </summary>
    public class XMLManager {
        private static readonly ILog log = LogManager.GetLogger(typeof(XMLManager));
        private static XNamespace ns = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync";

        public XMLManager() {
        }
        
        /// <summary>
        /// Exports any object given in "obj" to an xml file given in "filename"
        /// </summary>
        /// <param name="obj">The object that is to be serialized/exported to XML.</param>
        /// <param name="filename">The filename of the xml file to be written.</param>
        public static void Export(Object obj, string filename) {
            XmlTextWriter writer = new XmlTextWriter(filename, null);
            writer.Formatting = Formatting.Indented;
            writer.Indentation = 4;
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
                log.Error("Failed to import settings");
                OGCSexception.Analyse(ex);
                if (MainForm.Instance != null) MainForm.Instance.tabApp.SelectedTab = MainForm.Instance.tabPage_Settings;
            }
            fs.Close();
            return result;
        }

        public static string ImportElement(string nodeName, string filename) {
            try {
                XDocument xml = XDocument.Load(filename);
                XElement settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
                XElement xe = settingsXE.Elements(ns + nodeName).First();
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
            } catch (System.Exception ex) {
                log.Error("Failed retrieving '" + nodeName + "' from " + filename);
                OGCSexception.Analyse(ex);
                return null;
            }
        }
        
        public static void ExportElement(string nodeName, object nodeValue, string filename) {
            XDocument xml = XDocument.Load(filename);
            XElement settingsXE = xml.Descendants(ns + "Settings").FirstOrDefault();
            try {
                XElement xe = settingsXE.Elements(ns + nodeName).First();
                xe.SetValue(nodeValue);
                xml.Save(filename);
                log.Debug("Setting '" + nodeName + "' updated to '" + nodeValue + "'");
            } catch (Exception ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80131509") { //Sequence contains no elements
                    log.Debug("Adding Setting " + nodeName + " to settings.xml");
                    //This appends to the end, which won't import properly. 
                    //settingsXE.Add(new XElement(ns + nodeName, nodeValue));
                    //To save writing a sort method, let's just save everything!
                    Settings.Instance.Save();

                } else {
                    OGCSexception.Analyse(ex);
                    log.Error("Failed to export setting " + nodeName + "=" + nodeValue + " to settings.xml file.");
                }
            }
        }
    }
}
