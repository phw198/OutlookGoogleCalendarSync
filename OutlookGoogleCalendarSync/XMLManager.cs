using System;
using System.IO;
using System.Runtime.Serialization;
using System.Xml;

namespace OutlookGoogleCalendarSync
{
    /// <summary>
    /// Exports or imports any object to/from XML.
    /// </summary>
    public class XMLManager
    {
        public XMLManager()
        {
        }
        
        /// <summary>
        /// Exports any object given in "obj" to an xml file given in "filename"
        /// </summary>
        /// <param name="obj">The object that is to be serialized/exported to XML.</param>
        /// <param name="filename">The filename of the xml file to be written.</param>
        public static void export(Object obj, string filename)
        {
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
        public static T import<T>(string filename)
        {
            FileStream fs = new FileStream(filename, FileMode.Open);
            T result = default(T);
            try {
                result = (T)new DataContractSerializer(typeof(T)).ReadObject(fs);
            } catch {
                MainForm.Instance.tabSettings.SelectedTab = MainForm.Instance.tabPage_Settings;
            }
            fs.Close();
            return result;
        }
        
    }
}
