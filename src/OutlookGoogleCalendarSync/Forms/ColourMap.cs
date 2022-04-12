using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class ColourMap : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(ColourMap));
        public static Extensions.OutlookColourPicker OutlookComboBox = new Extensions.OutlookColourPicker();
        public static Extensions.GoogleColourPicker GoogleComboBox = new Extensions.GoogleColourPicker();
        
        public ColourMap() {
            log.Info("Opening colour mapping window.");
            OutlookComboBox = null;
            OutlookComboBox = new Extensions.OutlookColourPicker();
            OutlookComboBox.AddCategoryColours();
            GoogleComboBox = null;
            GoogleComboBox = new Extensions.GoogleColourPicker();
            GoogleComboBox.AddPaletteColours(true);

            InitializeComponent();
            loadConfig();
            OutlookOgcs.Calendar.Disconnect(true);
        }
        
        private void ColourMap_Shown(object sender, EventArgs e) {
            ddOutlookColour_SelectedIndexChanged(null, null);
        }

        private void loadConfig() {
            try {
                if (Forms.Main.Instance.ActiveCalendarProfile.ColourMaps.Count > 0) colourGridView.Rows.Clear();
                foreach (KeyValuePair<String, String> colourMap in Forms.Main.Instance.ActiveCalendarProfile.ColourMaps) {
                    addRow(colourMap.Key, GoogleOgcs.EventColour.Palette.GetColourName(colourMap.Value));
                }
                ddOutlookColour.AddCategoryColours();
                if (ddOutlookColour.Items.Count > 0)
                    ddOutlookColour.SelectedIndex = 0;

                ddGoogleColour.AddPaletteColours();
                if (ddGoogleColour.Items.Count > 0)
                    ddGoogleColour.SelectedIndex = 0;

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Populating gridview cells from Settings.", ex);
            }
        }

        private void addRow(String outlookColour, String googleColour) {
            int lastRow = 0;
            try {
                lastRow = colourGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentValue = colourGridView.Rows[lastRow].Cells["OutlookColour"].Value;
                if (currentValue != null && currentValue.ToString() != "") {
                    lastRow++;
                    colourGridView.Rows.Insert(lastRow);
                }
                colourGridView.Rows[lastRow].Cells["OutlookColour"].Value = outlookColour;
                colourGridView.Rows[lastRow].Cells["GoogleColour"].Value = googleColour;

                colourGridView.CurrentCell = colourGridView.Rows[lastRow].Cells[1];
                colourGridView.NotifyCurrentCellDirty(true);
                colourGridView.NotifyCurrentCellDirty(false);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Adding colour/category map row #" + lastRow, ex);
            }
        }

        private void newRowNeeded() {
            int lastRow = 0;
            try {
                lastRow = colourGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentOValue = colourGridView.Rows[lastRow].Cells["OutlookColour"].Value;
                Object currentGValue = colourGridView.Rows[lastRow].Cells["GoogleColour"].Value;
                if (currentOValue != null && currentOValue.ToString() != "" &&
                    currentGValue != null && currentGValue.ToString() != "") {
                    lastRow++;
                    DataGridViewCell lastCell = colourGridView.Rows[lastRow - 1].Cells[1];
                    if (lastCell != colourGridView.CurrentCell)
                        colourGridView.CurrentCell = lastCell;
                    colourGridView.NotifyCurrentCellDirty(true);
                    colourGridView.NotifyCurrentCellDirty(false);
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("newRowNeeded(): Adding colour/category map row #" + lastRow, ex);
            }
        }

        #region EVENTS
        private void btOK_Click(object sender, EventArgs e) {
            log.Fine("Checking no duplicate mappings exist.");
            SettingsStore.Calendar profile = Forms.Main.Instance.ActiveCalendarProfile;
            try {
                List<String> oColValues = new List<String>();
                List<String> gColValues = new List<String>();
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    oColValues.Add(row.Cells["OutlookColour"].Value.ToString());
                    gColValues.Add(row.Cells["GoogleColour"].Value.ToString());
                }
                String oDuplicates = string.Join("\r\n", oColValues.GroupBy(v => v).Where(g => g.Count() > 1).Select(s => "- "+ s.Key).ToList());
                String gDuplicates = string.Join("\r\n", gColValues.GroupBy(v => v).Where(g => g.Count() > 1).Select(s => "- " + s.Key).ToList());

                if (!string.IsNullOrEmpty(oDuplicates) && (profile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id || profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)) {
                    OgcsMessageBox.Show("The following Outlook categories cannot be mapped more than once:-\r\n\r\n" + oDuplicates, "Duplicate Outlook Mappings", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                } else if (!string.IsNullOrEmpty(gDuplicates) && (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id || profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)) {
                    OgcsMessageBox.Show("The following Google colours cannot be mapped more than once:-\r\n\r\n" + gDuplicates, "Duplicate Google Mappings", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed looking for duplicating mappings before storing in Settings.", ex);
                OgcsMessageBox.Show("An error was encountered storing your custom mappings.", "Cannot save mappings", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try {
                log.Fine("Storing colour mappings in Settings.");
                profile.ColourMaps.Clear();
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    if (string.IsNullOrEmpty(row.Cells["OutlookColour"].Value?.ToString()?.Trim()) || string.IsNullOrEmpty(row.Cells["GoogleColour"].Value?.ToString()?.Trim())) continue;
                    try {
                        profile.ColourMaps.Add(row.Cells["OutlookColour"].Value.ToString(), GoogleOgcs.EventColour.Palette.GetColourId(row.Cells["GoogleColour"].Value.ToString()));
                    } catch (System.ArgumentException ex) {
                        if (OGCSexception.GetErrorCode(ex) == "0x80070057") {
                            //An item with the same key has already been added
                        } else throw;
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not save colour/category mappings to Settings.", ex);
            } finally {
                this.Close();
            }
        }

        private void colourGridView_CellClick(object sender, DataGridViewCellEventArgs e) {
            if (!this.Visible) return;

            Boolean validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            //Check to make sure the cell clicked is the cell containing the combobox 
            if (validClick && colourGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) {
                colourGridView.BeginEdit(true);
                ((ComboBox)colourGridView.EditingControl).DroppedDown = true;
            }
        }
        
        private void colourGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e) {
            try {
                if (e.Control is ComboBox) {
                    ComboBox cb = e.Control as ComboBox;
                    cb.DrawMode = DrawMode.OwnerDrawFixed;
                    if (cb is Extensions.OutlookColourCombobox) {
                        cb.DrawItem -= OutlookComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += OutlookComboBox.ColourPicker_DrawItem;
                        OutlookComboBox.ColourPicker_DrawItem(sender, null);
                    } else if (cb is Extensions.GoogleColourCombobox) {
                        cb.DrawItem -= GoogleComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += GoogleComboBox.ColourPicker_DrawItem;
                        GoogleComboBox.ColourPicker_DrawItem(sender, null);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private void colourGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e) {
            newRowNeeded();
        }

        private void colourGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
            if (!this.Visible) return;
            
            if (colourGridView.CurrentCell.ColumnIndex == 0)
                ddGoogleColour_SelectedIndexChanged(null, null);
            else if (colourGridView.CurrentCell.ColumnIndex == 1)
                ddOutlookColour_SelectedIndexChanged(null, null);
        }

        private void colourGridView_CellEnter(object sender, DataGridViewCellEventArgs e) {
            if (colourGridView.CurrentRow.Index + 1 < colourGridView.Rows.Count) return;

            newRowNeeded();
        }

        private void colourGridView_SelectionChanged(object sender, EventArgs e) {
            //Protect against the last row being selected for deletion
            try {
                if (colourGridView.SelectedRows.Count == 0) return;

                int selectedRow = colourGridView.SelectedRows[colourGridView.SelectedRows.Count - 1].Index;
                if (selectedRow == colourGridView.Rows.Count - 1) {
                    log.Debug("Last row");
                    colourGridView.Rows[selectedRow].Selected = false;
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private void ddOutlookColour_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            try {
                ddGoogleColour.SelectedIndexChanged -= ddGoogleColour_SelectedIndexChanged;
                
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    if (row.Cells["OutlookColour"].Value.ToString() == ddOutlookColour.SelectedItem.Text && !string.IsNullOrEmpty(row.Cells["GoogleColour"].Value?.ToString())) {
                        String colourId = GoogleOgcs.EventColour.Palette.GetColourId(row.Cells["GoogleColour"].Value.ToString());
                        ddGoogleColour.SelectedIndex = Convert.ToInt16(colourId);
                        return;
                    }
                }

                ddGoogleColour.SelectedIndex = Convert.ToInt16(GoogleOgcs.Calendar.Instance.GetColour(ddOutlookColour.SelectedItem.OutlookCategory).Id);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("ddOutlookColour_SelectedIndexChanged(): Could not update ddGoogleColour.", ex);
            } finally {
                ddGoogleColour.SelectedIndexChanged += ddGoogleColour_SelectedIndexChanged;
            }
        }
        private void ddGoogleColour_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            try {
                ddOutlookColour.SelectedIndexChanged -= ddOutlookColour_SelectedIndexChanged;

                String oCatName = null;
                log.Fine("Checking grid for map...");
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    if (row.Cells["GoogleColour"].Value != null && row.Cells["GoogleColour"].Value.ToString() == ddGoogleColour.SelectedItem.Name) {
                        oCatName = row.Cells["OutlookColour"].Value.ToString();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(oCatName))
                    oCatName = OutlookOgcs.Calendar.Instance.GetCategoryColour(ddGoogleColour.SelectedItem.Id, false);

                if (!string.IsNullOrEmpty(oCatName)) {
                    foreach (OutlookOgcs.Categories.ColourInfo cInfo in ddOutlookColour.Items) {
                        if (cInfo.Text == oCatName) {
                            ddOutlookColour.SelectedItem = cInfo;
                            return;
                        }
                    }
                    log.Warn("The category '" + oCatName + "' exists, but wasn't found in Outlook colour dropdown.");
                    OutlookOgcs.Calendar.Instance.IOutlook.RefreshCategories();
                    while (ddOutlookColour.Items.Count > 0)
                        ddOutlookColour.Items.RemoveAt(0);
                    ddOutlookColour.AddCategoryColours();

                    foreach (OutlookOgcs.Categories.ColourInfo cInfo in ddOutlookColour.Items) {
                        if (cInfo.Text == oCatName) {
                            ddOutlookColour.SelectedItem = cInfo;
                            return;
                        }
                    }
                }
                ddOutlookColour.SelectedIndex = 0;

            } catch (System.Exception ex) {
                OGCSexception.Analyse("ddGoogleColour_SelectedIndexChanged(): Could not update ddOutlookColour.", ex);
            } finally {
                ddOutlookColour.SelectedIndexChanged += ddOutlookColour_SelectedIndexChanged;
            }
        }
        #endregion
    }
}
