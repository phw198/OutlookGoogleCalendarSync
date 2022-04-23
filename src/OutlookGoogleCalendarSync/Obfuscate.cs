using log4net;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    [DataContract]
    public class Obfuscate {
        private static readonly ILog log = LogManager.GetLogger(typeof(Obfuscate));
        private const int findCol = 0;
        private const int replaceCol = 1;
            
        public Obfuscate() {
            setDefaults();
        }

        [DataMember] public bool Enabled { get; set; }
        [DataMember] public Sync.Direction Direction { get; set; }
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
            this.Direction = Sync.Direction.OutlookToGoogle;
            this.findReplace = new List<FindReplace>();
        }

        public void SaveRegex(DataGridView data) {
            findReplace = new List<FindReplace>();

            if (data.IsCurrentCellDirty) {
                data.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

            foreach (DataGridViewRow row in data.Rows) {
                if (row.Cells[findCol].Value != null) {
                    this.findReplace.Add(
                        new FindReplace(
                            row.Cells[findCol].Value.ToString(), 
                            row.Cells[replaceCol].Value == null ? "" : row.Cells[replaceCol].Value.ToString() 
                        ));
                    
                }
            }
        }

        public void LoadRegex(DataGridView data) {
            int dataRow = 0;
            foreach (FindReplace regex in findReplace) {
                data.Rows[dataRow].Cells[findCol].Value = regex.find;
                data.Rows[dataRow].Cells[replaceCol].Value = regex.replace;
                data.CurrentCell = data.Rows[dataRow].Cells[0];
                data.NotifyCurrentCellDirty(true);
                data.NotifyCurrentCellDirty(false);
                dataRow++;
            }
            data.CurrentCell = data.Rows[0].Cells[0];
        }

        /// <summary>
        /// Apply regular expression to a source string, if syncing in the correct direction
        /// </summary>
        /// <param name="source">Source calendar string</param>
        /// <param name="target">Target calendar string, null if creating</param>
        /// <param name="direction">Direction engine is being synced</param>
        /// <returns></returns>
        public static String ApplyRegex(String source, String target, Sync.Direction direction) {
            String retStr = source ?? "";
            if (Sync.Engine.Calendar.Instance.Profile.Obfuscation.Enabled) {
                if (direction.Id == Sync.Engine.Calendar.Instance.Profile.Obfuscation.Direction.Id) {
                    foreach (DataGridViewRow row in Forms.Main.Instance.dgObfuscateRegex.Rows) {
                        DataGridViewCellCollection cells = row.Cells;
                        if (cells[Obfuscate.findCol].Value != null) {
                            System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(
                                cells[Obfuscate.findCol].Value.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                            System.Text.RegularExpressions.MatchCollection matches = rgx.Matches(retStr);
                            if (matches.Count > 0) {
                                log.Fine("Regex has matched and altered string: " + cells[Obfuscate.findCol].Value.ToString());
                                if (cells[Obfuscate.replaceCol].Value == null) cells[Obfuscate.replaceCol].Value = "";
                                retStr = rgx.Replace(retStr, cells[Obfuscate.replaceCol].Value.ToString());
                            }
                        }
                    }
                } else {
                    retStr = target ?? source;
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
