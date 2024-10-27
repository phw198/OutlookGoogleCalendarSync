using log4net;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    [DataContract]
    public class Obfuscate {
        private static readonly ILog log = LogManager.GetLogger(typeof(Obfuscate));

        public enum Columns : int {
            find = 0,
            replace = 1,
            target = 2
        }
        public enum Property {
            Description,
            Location,
            Subject
        }

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
                this.findReplace.Clear();
                foreach (FindReplace fr in value) {
                    this.findReplace.Add(new FindReplace(fr.find, fr.replace, fr.target));
                }
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
            this.findReplace = new List<FindReplace> { new FindReplace(null, null, "S") } ;
        }

        public void SaveRegex(DataGridView data) {
            findReplace = new List<FindReplace>();

            if (data.IsCurrentCellDirty) {
                data.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

            foreach (DataGridViewRow row in data.Rows) {
                if (String.IsNullOrEmpty(row.Cells[((int)Columns.find)].Value?.ToString())) continue;

                this.findReplace.Add(
                    new FindReplace(
                        row.Cells[((int)Columns.find)].Value.ToString(),
                        row.Cells[((int)Columns.replace)].Value == null ? "" : row.Cells[((int)Columns.replace)].Value.ToString(),
                        String.IsNullOrEmpty(row.Cells[((int)Columns.target)].Value?.ToString().Trim()) ? "S" : row.Cells[((int)Columns.target)].Value.ToString()
                    ));
            }
        }

        public void LoadRegex(DataGridView data) {
            int dataRow = 0;
            foreach (FindReplace regex in findReplace) {
                if (String.IsNullOrEmpty(regex.find)) continue;

                data.Rows[dataRow].Cells[((int)Columns.find)].Value = regex.find;
                data.Rows[dataRow].Cells[((int)Columns.replace)].Value = regex.replace;
                data.Rows[dataRow].Cells[((int)Columns.target)].Value = String.IsNullOrEmpty(regex.target) ? "S" : regex.target;
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
        public static String ApplyRegex(Property property, String source, String target, Sync.Direction direction) {
            String retStr = source ?? "";
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (profile.Obfuscation.Enabled) {
                if (direction.Id == profile.Obfuscation.Direction.Id) {
                    foreach (FindReplace regex in profile.Obfuscation.FindReplace) {
                        if (String.IsNullOrEmpty(regex.find)) continue;

                        if (property == Property.Subject && regex.target.Contains("S") ||
                            property == Property.Location && regex.target.Contains("L") ||
                            property == Property.Description && regex.target.Contains("D")
                        ) {
                            System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(
                                regex.find, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                            System.Text.RegularExpressions.MatchCollection matches = rgx.Matches(retStr);
                            if (matches.Count > 0) {
                                log.Fine("Regex has matched and altered "+ property.ToString() +" string: " + regex.find);
                                retStr = rgx.Replace(retStr, regex.replace ?? "");
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
        public FindReplace(String find, String replace, String target) {
            this.find = find;
            this.replace = replace;
            this.target = target ?? "S";
        }

        [DataMember]
        public string find { get; internal set; }
        [DataMember]
        public string replace { get; internal set; }
        [DataMember]
        public string target { get; internal set; }
    }
}
