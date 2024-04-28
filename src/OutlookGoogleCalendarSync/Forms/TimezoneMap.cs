using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class TimezoneMap : Form {

        private static readonly ILog log = LogManager.GetLogger(typeof(TimezoneMap));
        private const string tzMapFile = "tzmap.xml";

        public static TimeZoneInfo TimezoneMap_StaThread(String organiserTz, TimeZoneInfo bestEffortTzi, System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ) {
            System.Threading.Thread tzThread = new System.Threading.Thread(() => {
                TimezoneMap tzMap = new TimezoneMap(organiserTz, bestEffortTzi);
                //This is "safe" in that this form won't access the Main form thread - and if it does we can do that bit threadsafe
                //More important to get this form displaying on top of the Main form though, hence this hack
                Control.CheckForIllegalCrossThreadCalls = false;
                tzMap.ShowDialog(Forms.Main.Instance);
                Control.CheckForIllegalCrossThreadCalls = true;
            });
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
                col.DataSource = new BindingSource(cbTz, null);
                col.DisplayMember = "Value";
                col.ValueMember = "Key";
                col.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                
                tzGridView.Columns.RemoveAt(1);
                tzGridView.Columns.Add(col);

                loadConfig();

            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }

        private void loadConfig() {
            try {
                tzGridView.AllowUserToAddRows = true;
                if (Settings.Instance.TimezoneMaps.Count > 0) tzGridView.Rows.Clear();
                foreach (KeyValuePair<String, String> tzMap in Settings.Instance.TimezoneMaps) {
                    addRow(tzMap.Key, tzMap.Value);
                }

            } catch (System.Exception ex) {
                ex.Analyse("Populating gridview cells from Settings.");
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
                ex.Analyse("Adding timezone map row #" + lastRow);
            }
        }

        public static TimeZoneInfo GetSystemTimezone(String organiserTz, System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ) {
            TimeZoneInfo tzi = null;
            if (Settings.Instance.TimezoneMaps.ContainsKey(organiserTz)) {
                tzi = sysTZ.FirstOrDefault(t => t.Id == Settings.Instance.TimezoneMaps[organiserTz]);
                if (tzi != null) {
                    log.Debug("Using custom timezone mapping ID '" + tzi.Id + "' for '" + organiserTz + "'");
                    return tzi;
                } else log.Warn("Failed to convert custom timezone mapping to any available system timezone.");
            }
            return tzi;
        }

        #region EVENTS
        private void btOK_Click(object sender, EventArgs e) {
            try {
                Settings.Instance.TimezoneMaps.Clear();
                foreach (DataGridViewRow row in tzGridView.Rows) {
                    if (row.Cells[0].Value == null || row.Cells[0].Value.ToString().Trim() == "") continue;
                    try {
                        Settings.Instance.TimezoneMaps.Add(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString());
                    } catch (System.ArgumentException ex) {
                        if (ex.GetErrorCode() == "0x80070057") {
                            //An item with the same key has already been added
                        } else throw;
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not save timezone mappings to Settings.");
            } finally {
                this.Close();
                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.btCustomTzMap, "Visible", Settings.Instance.TimezoneMaps.Count != 0);
                Settings.Instance.Save();
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
                    e.Exception.Analyse("Bad cell value in timezone data grid.");
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
