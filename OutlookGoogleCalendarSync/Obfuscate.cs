using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using log4net;
using System.Windows.Forms;

//[assembly: ContractNamespaceAttribute("http://www.cohowinery.com/employees",
//    ClrNamespace = "OutlookGoogleCalendarSync")]

namespace OutlookGoogleCalendarSync {
    [DataContract]
    public class Obfuscate {
        private static readonly ILog log = LogManager.GetLogger(typeof(Obfuscate));
        public const int FindCol = 0;
        public const int ReplaceCol = 1;
            
        public Obfuscate() {
            setDefaults();
        }

        [DataMember] public bool Enabled { get; set; }
        [DataMember] public SyncDirection Direction { get; set; }
        private List<FindReplace> findReplace;
        
        [DataMember] public List<FindReplace> FindReplace {
            get {
                return this.findReplace ?? new List<FindReplace>();
            }
            set {
                this.findReplace = value;
            }
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }
        
        private void setDefaults() {
            this.Enabled = false;
            this.Direction = SyncDirection.OutlookToGoogle;
            this.findReplace = new List<FindReplace>();
        }

        public void SaveRegex(DataGridView data) {
            findReplace = new List<FindReplace>();

            foreach (DataGridViewRow row in data.Rows) {
                if (row.Cells[FindCol].Value != null) {
                    this.findReplace.Add(
                        new FindReplace(
                            row.Cells[FindCol].Value.ToString(), 
                            row.Cells[ReplaceCol].Value == null ? "" : row.Cells[ReplaceCol].Value.ToString() 
                        ));
                    
                }
            }
        }

        public void LoadRegex(DataGridView data) {
            int dataRow = 0;
            foreach (FindReplace regex in findReplace) {
                data.Rows[dataRow].Cells[FindCol].Value = regex.find;
                data.Rows[dataRow].Cells[ReplaceCol].Value = regex.replace;
                data.CurrentCell = data.Rows[dataRow].Cells[0];
                data.NotifyCurrentCellDirty(true);
                data.NotifyCurrentCellDirty(false);
                dataRow++;
            }
            data.CurrentCell = data.Rows[0].Cells[0];
        }

        public static String ApplyRegex(String source, SyncDirection direction) {
            String retStr = source;
            if (Settings.Instance.Obfuscation.Enabled && direction == Settings.Instance.Obfuscation.Direction) {
                foreach (DataGridViewRow row in MainForm.Instance.dgObfuscateRegex.Rows) {
                    DataGridViewCellCollection cells = row.Cells;
                    if (cells[Obfuscate.FindCol].Value != null) {
                        System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(
                            cells[Obfuscate.FindCol].Value.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        System.Text.RegularExpressions.MatchCollection matches = rgx.Matches(retStr);
                        if (matches.Count > 0) {
                            log.Debug("Regex has matched and altered string: " + cells[Obfuscate.FindCol].Value.ToString());
                            if (cells[Obfuscate.ReplaceCol].Value == null) cells[Obfuscate.ReplaceCol].Value = "";
                            retStr = rgx.Replace(retStr, cells[Obfuscate.ReplaceCol].Value.ToString());
                        }
                    }
                }
            }
            return retStr;
        }
    }

    [DataContract]
    public class FindReplace {
        public FindReplace(String find, String replace) {
            this.find = find;
            this.replace = replace;
        }

        [DataMember]
        public string find { get; set; }
        [DataMember]
        public string replace { get; set; }
    }
}
