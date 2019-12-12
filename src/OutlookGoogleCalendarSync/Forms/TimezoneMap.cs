using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class TimezoneMap : Form {

        private static readonly ILog log = LogManager.GetLogger(typeof(TimezoneMap));
        private const string tzMapFile = "tzmap.xml";

        public static TimeZoneInfo TimezoneMap_StaThread(String organiserTz, TimeZoneInfo bestEffortTzi, System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ) {
            System.Threading.Thread tzThread = new System.Threading.Thread(() => new TimezoneMap(organiserTz, bestEffortTzi).ShowDialog());
            tzThread.SetApartmentState(System.Threading.ApartmentState.STA);
            tzThread.Start();
            tzThread.Join();
            return GetSystemTimezone(organiserTz, sysTZ);
        }

        public TimezoneMap() {
            InitializeComponent();
            initialiseDataGridView();
            tzGridView.AllowUserToAddRows = false;
        }
        private TimezoneMap(String organiserTz, TimeZoneInfo organiserTzi) {
            InitializeComponent();
            initialiseDataGridView();
            addRow(organiserTz, organiserTzi.Id);
            tzGridView.AllowUserToAddRows = false;
        }

        private void initialiseDataGridView() {
            try {
                log.Info("Opening timezone mapping window.");
                
                log.Fine("Building default system timezone dropdowns.");
                System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ = TimeZoneInfo.GetSystemTimeZones();
                Dictionary<String, String> cbTz = new Dictionary<String, String>();
                sysTZ.ToList().ForEach(tzi => cbTz.Add(tzi.Id, tzi.DisplayName));

                //Replace existing TZ column with custom dropdown
                DataGridViewComboBoxColumn col = tzGridView.Columns[1] as DataGridViewComboBoxColumn;
                col.DataSource = new BindingSource(cbTz, null); //bs;
                col.DisplayMember = "Value";
                col.ValueMember = "Key";
                col.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                
                tzGridView.Columns.RemoveAt(1);
                tzGridView.Columns.Add(col);

                loadConfig();

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private void loadConfig() {
            try {
                tzGridView.AllowUserToAddRows = true;
                if (Settings.Instance.TimezoneMapping.Count > 0) tzGridView.Rows.Clear();
                foreach (KeyValuePair<String, String> tzMap in Settings.Instance.TimezoneMapping) {
                    addRow(tzMap.Key, tzMap.Value);
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Populating gridview cells from Settings.", ex);
            }
        }

        private void addRow(String organiserTz, String systemTz) {
            int lastRow = 0;
            try {
                lastRow = tzGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentValue = tzGridView.Rows[lastRow].Cells["OrganiserTz"].Value;
                if (currentValue != null && currentValue.ToString() != "") {
                    lastRow++;
                    tzGridView.Rows.Insert(lastRow);
                }
                tzGridView.Rows[lastRow].Cells["OrganiserTz"].Value = organiserTz;
                tzGridView.Rows[lastRow].Cells["SystemTz"].Value = systemTz;

                tzGridView.CurrentCell = tzGridView.Rows[lastRow].Cells[1];
                tzGridView.NotifyCurrentCellDirty(true);
                tzGridView.NotifyCurrentCellDirty(false);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Adding timezone map row #" + lastRow, ex);
            }
        }

        public static Dictionary<String, String> LoadConfigFromXml() {
            Dictionary<String, String> config = new Dictionary<String, String>();

            try {
                String xmlFile = Path.Combine(Program.UserFilePath, tzMapFile);
                if (!File.Exists(xmlFile)) return config;

                log.Debug("Loading timezone mappings from " + tzMapFile);
                XmlDocument xmlDoc = new XmlDocument();
                FileStream fs = new FileStream(xmlFile, FileMode.Open, FileAccess.Read);
                try {
                    xmlDoc.Load(fs);
                    XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/TimeZoneMaps/TimeZoneMap");
                    foreach (XmlNode node in nodeList) {
                        config.Add(node.SelectSingleNode("OrganiserTz").InnerText, node.SelectSingleNode("SystemTz").InnerText);
                    }
                } catch (System.Xml.XmlException ex) {
                    if (OGCSexception.GetErrorCode(ex) == "0x80131940") { //Root element is missing.
                        log.Debug(tzMapFile + " is empty.");
                    } else
                        throw;
                } finally {
                    fs.Close();
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
            return config;
        }

        public static TimeZoneInfo GetSystemTimezone(String organiserTz, System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ) {
            TimeZoneInfo tzi = null;
            if (Settings.Instance.TimezoneMapping.ContainsKey(organiserTz)) {
                tzi = sysTZ.FirstOrDefault(t => t.Id == Settings.Instance.TimezoneMapping[organiserTz]);
                if (tzi != null) {
                    log.Debug("Using custom timezone mapping ID '" + tzi.Id + "' for '" + organiserTz + "'");
                    return tzi;
                } else log.Warn("Failed to convert custom timezone mapping to any available system timezone.");
            }
            return tzi;
        }

        #region EVENTS
        private void btSave_Click(object sender, EventArgs e) {
            try {
                //Building dataTable
                DataTable dt = new DataTable();
                dt.TableName = "TimeZoneMap";

                foreach (DataGridViewColumn col in tzGridView.Columns) {
                    DataColumn dc = new DataColumn(col.Name);
                    dt.Columns.Add(dc);
                }
                foreach (DataGridViewRow row in tzGridView.Rows) {
                    if (row.Cells[0].Value == null || row.Cells[0].Value.ToString().Trim() == "") continue;
                    DataRow dr = dt.NewRow();
                    dr[0] = row.Cells[0].Value;
                    dr[1] = row.Cells[1].Value;
                    dt.Rows.Add(dr);
                }
                DataSet ds = new DataSet();
                ds.DataSetName = "TimeZoneMaps";
                ds.Tables.Add(dt);

                XmlTextWriter xmlSave = null;
                try {
                    xmlSave = new XmlTextWriter(System.IO.Path.Combine(Program.UserFilePath, tzMapFile), Encoding.UTF8);
                    xmlSave.Formatting = Formatting.Indented;
                    xmlSave.Indentation = 4;
                    ds.WriteXml(xmlSave);
                } finally {
                    xmlSave.Close();
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not save timezone mappings to XML.", ex);
            } finally {
                Settings.LoadTimezoneMap();
                this.Close();
            }
        }

        private void tzGridView_DataError(object sender, DataGridViewDataErrorEventArgs e) {
            log.Error(e.Context.ToString());
            if (e.Exception.HResult == -2147024809) { //DataGridViewComboBoxCell value is not valid.
                DataGridViewCell cell = tzGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                log.Warn("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] has invalid value of '" + cell.Value + "'. Removing.");
                cell.OwningRow.Cells[0].Value = null;
                cell.OwningRow.Cells[1].Value = null;
            } else {
                try {
                    DataGridViewCell cell = tzGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    log.Debug("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] caused error.");
                } catch {
                } finally {
                    OGCSexception.Analyse("Bad cell value in timezone data grid.", e.Exception);
                }
            }
        }

        private void tzGridView_CellClick(object sender, DataGridViewCellEventArgs e) {
            if (!this.Visible) return;

            Boolean validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            //Check to make sure the cell clicked is the cell containing the combobox 
            if (validClick && tzGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) {
                tzGridView.BeginEdit(true);
                ((ComboBox)tzGridView.EditingControl).DroppedDown = true;
            }
        }
        #endregion
    }
}
