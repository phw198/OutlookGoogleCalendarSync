using Ogcs = OutlookGoogleCalendarSync;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Extensions {
    #region Dropdown colour pickers
    public abstract class ColourPicker : ComboBox {
        public ColourPicker() {
            DropDownStyle = ComboBoxStyle.DropDownList;
            DrawMode = DrawMode.OwnerDrawFixed;
        }
    }

    public class OutlookColourPicker : ColourPicker {
        public OutlookColourPicker() {
            base.DrawItem += ColourPicker_DrawItem;
        }
        public void AddColourItems() {
            Items.Clear();
            AddCategoryColours();
        }

        /// <summary>
        /// Add all the available Outlook colours
        /// </summary>
        public void AddStandardColours() {
            foreach (KeyValuePair<OlCategoryColor, Color> colour in Outlook.Categories.Map.Colours) {
                Items.Add(new Outlook.Categories.ColourInfo(colour.Key, colour.Value));
            }
        }

        /// <summary>
        /// Add just the colours associated with categories
        /// </summary>
        public void AddCategoryColours() {
            if (Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003) return;

            Items.AddRange(Outlook.Calendar.Categories.DropdownItems().ToArray());
        }

        public void ColourPicker_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            if (e == null || e.Index < 0 || e.Index >= cbColour.Items.Count)
                return;

            // Get the colour
            Outlook.Categories.ColourInfo colour = (Outlook.Categories.ColourInfo)Items[e.Index];
            ColourCombobox.DrawComboboxItemColour(cbColour, new SolidBrush(colour.Colour), colour.Text, e);
        }

        public new Outlook.Categories.ColourInfo SelectedItem {
            get { return (Outlook.Categories.ColourInfo)base.SelectedItem; }
            set { base.SelectedItem = value; }
        }
    }

    public class GoogleColourPicker : ColourPicker {
        public GoogleColourPicker() {
            DrawItem += ColourPicker_DrawItem;
            Enter += ColourPicker_Enter;
        }

        /// <summary>
        /// Add all the available Google colours
        /// </summary>
        public void AddPaletteColours(Boolean connectToGoogle = false) {
            if (Ogcs.Google.Calendar.IsInstanceNull && !connectToGoogle) return;

            foreach (Ogcs.Google.EventColour.Palette palette in Ogcs.Google.Calendar.Instance.ColourPalette.ActivePalette) {
                Items.Add(palette);
            }
        }

        public void ColourPicker_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            if (e == null || e.Index < 0 || e.Index >= cbColour.Items.Count)
                return;

            // Get the colour
            Ogcs.Google.EventColour.Palette colour = (Ogcs.Google.EventColour.Palette)Items[e.Index];
            ColourCombobox.DrawComboboxItemColour(cbColour, new SolidBrush(colour.RgbValue), colour.Name, e);
        }

        public new Ogcs.Google.EventColour.Palette SelectedItem {
            get { return (Ogcs.Google.EventColour.Palette)base.SelectedItem; }
            set { base.SelectedItem = value; }
        }
        
        private void ColourPicker_Enter(object sender, EventArgs e) {
            if (Forms.Main.Instance.ActiveCalendarProfile.UseGoogleCalendar == null || string.IsNullOrEmpty(Forms.Main.Instance.ActiveCalendarProfile.UseGoogleCalendar.Id)) {
                Ogcs.Extensions.MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.", "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ToolTip loading = new ToolTip();
            try {
                Ogcs.Google.EventColour.Palette currentSelection = null;

                if (Ogcs.Google.Calendar.IsInstanceNull || !Ogcs.Google.Calendar.Instance.ColourPalette.IsCached()) {
                    loading.SetToolTip(this, "Retrieving colours from Google...");
                    loading.ShowAlways = true;
                    loading.InitialDelay = 0;
                    loading.Show("Retrieving colours from Google...", this, this.FindForm().PointToClient(this.Parent.PointToScreen(this.Location)));

                    Ogcs.Google.Calendar.Instance.ColourPalette.Get();
                    currentSelection = (Ogcs.Google.EventColour.Palette)SelectedItem;
                    
                    loading.Hide(this);
                }
                if (Items.Count != Ogcs.Google.Calendar.Instance.ColourPalette.ActivePalette.Count) {
                    while (Items.Count > 0)
                        Items.RemoveAt(0);
                    AddPaletteColours(true);
                }

                foreach (Ogcs.Google.EventColour.Palette pInfo in Items) {
                    if (pInfo.Id == currentSelection?.Id) {
                        SelectedItem = pInfo;
                        break;
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("ColourPicker_Enter()");
            } finally {
                loading.Hide(this);
                loading.RemoveAll();
            }

            if (Items.Count > 1 && SelectedIndex == -1)
                SelectedIndex = 0;
        }
    }
    #endregion

    #region ColourComboboxColumns
    public class DataGridViewOutlookColourComboBoxColumn : DataGridViewColumn {
        public DataGridViewOutlookColourComboBoxColumn() : base(new DataGridViewOutlookColourComboBoxCell()) {
        }

        public override DataGridViewCell CellTemplate {
            get {
                return base.CellTemplate;
            }
            set {
                // Ensure that the cell used for the template is a DataGridViewOutlookColourComboBoxCell.
                if (value != null && !value.GetType().IsAssignableFrom(typeof(DataGridViewOutlookColourComboBoxCell))) {
                    throw new InvalidCastException("Must be a DataGridViewOutlookColourComboBoxCell");
                }
                base.CellTemplate = value;
            }
        }
    }

    public class DataGridViewGoogleColourComboBoxColumn : DataGridViewColumn {
        public DataGridViewGoogleColourComboBoxColumn() : base(new DataGridViewGoogleColourComboBoxCell()) {
        }

        public override DataGridViewCell CellTemplate {
            get {
                return base.CellTemplate;
            }
            set {
                // Ensure that the cell used for the template is a DataGridViewGoogleColourComboBoxCell.
                if (value != null && !value.GetType().IsAssignableFrom(typeof(DataGridViewGoogleColourComboBoxCell))) {
                    throw new InvalidCastException("Must be a DataGridViewGoogleColourComboBoxCell");
                }
                base.CellTemplate = value;
            }
        }
    }
    #endregion

    #region ColourComboBoxCells
    public class DataGridViewOutlookColourComboBoxCell : DataGridViewTextBoxCell {
        public DataGridViewOutlookColourComboBoxCell() : base() { }

        public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle) {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            OutlookColourCombobox ctl = DataGridView.EditingControl as OutlookColourCombobox;
            if (this.RowIndex >= 0) {
                if (this.Value == null)
                    ctl.SelectedIndex = 0;
                else {
                    String currentText = this.Value.ToString();
                    if (ctl.Items.Count == 0)
                        ctl.PopulateDropdownItems();
                    if (!string.IsNullOrEmpty(currentText)) this.Value = currentText;

                    foreach (Outlook.Categories.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                        if (ci.Text == (String)this.Value) {
                            ctl.SelectedValue = ci;
                            break;
                        }
                    }
                }
            }
        }

        public override Type EditType {
            get {
                return typeof(OutlookColourCombobox);
            }
        }

        public override Type ValueType {
            get {
                return typeof(Outlook.Categories.ColourInfo);
            }
        }

        public override object DefaultNewRowValue {
            get {
                return String.Empty;
            }
        }

        protected override void Paint(System.Drawing.Graphics graphics, System.Drawing.Rectangle clipBounds, System.Drawing.Rectangle cellBounds, int rowIndex, System.Windows.Forms.DataGridViewElementStates elementState, object value, object formattedValue, string errorText, System.Windows.Forms.DataGridViewCellStyle cellStyle, System.Windows.Forms.DataGridViewAdvancedBorderStyle advancedBorderStyle, System.Windows.Forms.DataGridViewPaintParts paintParts) {
            //Paint inactive cells
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, "", "", errorText, cellStyle, advancedBorderStyle, paintParts);

            try {
                if (rowIndex < 0)
                    return;

                foreach (Outlook.Categories.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                    if (ci.Text == this.Value.ToString()) {
                        Brush boxBrush = new SolidBrush(ci.Colour);
                        Brush textBrush = SystemBrushes.WindowText;
                        Extensions.ColourCombobox.DrawComboboxItemColour(true, boxBrush, textBrush, this.Value.ToString(), graphics, cellBounds);
                        break;
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("OutlookColourComboboxCell.Paint()");
            }
        }
    }

    public class DataGridViewGoogleColourComboBoxCell : DataGridViewTextBoxCell {
        public DataGridViewGoogleColourComboBoxCell() : base() { }

        public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle) {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            GoogleColourCombobox ctl = DataGridView.EditingControl as GoogleColourCombobox;
            if (this.RowIndex >= 0) {
                if (this.Value == null)
                    ctl.SelectedIndex = 0;
                else {
                    String currentText = this.Value.ToString();
                    if (ctl.Items.Count == 0)
                        ctl.PopulateDropdownItems();
                    if (!string.IsNullOrEmpty(currentText)) this.Value = currentText;

                    foreach (Ogcs.Google.EventColour.Palette p in Forms.ColourMap.GoogleComboBox.Items) {
                        if (p.Name == (String)this.Value) {
                            ctl.SelectedValue = p;
                            break;
                        }
                    }
                }
            }
        }

        public override Type EditType {
            get {
                return typeof(GoogleColourCombobox);
            }
        }

        public override Type ValueType {
            get {
                return typeof(Ogcs.Google.EventColour.Palette);
            }
        }

        public override object DefaultNewRowValue {
            get {
                return String.Empty;
            }
        }

        protected override void Paint(System.Drawing.Graphics graphics, System.Drawing.Rectangle clipBounds, System.Drawing.Rectangle cellBounds, int rowIndex, System.Windows.Forms.DataGridViewElementStates elementState, object value, object formattedValue, string errorText, System.Windows.Forms.DataGridViewCellStyle cellStyle, System.Windows.Forms.DataGridViewAdvancedBorderStyle advancedBorderStyle, System.Windows.Forms.DataGridViewPaintParts paintParts) {
            //Paint inactive cells
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, "", "", errorText, cellStyle, advancedBorderStyle, paintParts);

            try {
                if (rowIndex < 0)
                    return;

                foreach (Ogcs.Google.EventColour.Palette ci in Forms.ColourMap.GoogleComboBox.Items) {
                    if (ci.Name == this.Value?.ToString()) {
                        Brush boxBrush = new SolidBrush(ci.RgbValue);
                        Brush textBrush = SystemBrushes.WindowText;
                        Extensions.ColourCombobox.DrawComboboxItemColour(true, boxBrush, textBrush, this.Value.ToString(), graphics, cellBounds);
                        break;
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("GoogleColourComboboxCell.Paint()");
            }
        }
    }
    #endregion

    #region ColourComboboxes
    public abstract class ColourCombobox : ComboBox {
        DataGridView dataGridView;
        private bool valueChanged = false;
        int rowIndex;

        public object EditingControlFormattedValue {
            get {
                return this.FormatString;
            }
            set {
                if (value is String) {
                    try {
                        this.FormatString = (string)value;
                    } catch {
                        this.FormatString = string.Empty;
                    }
                }
            }
        }

        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) {
            return EditingControlFormattedValue;
        }

        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle) {
            this.Font = dataGridViewCellStyle.Font;
            this.ForeColor = dataGridViewCellStyle.ForeColor;
            this.BackColor = dataGridViewCellStyle.BackColor;
        }

        public DataGridView EditingControlDataGridView {
            get {
                return dataGridView;
            }
            set {
                dataGridView = value;
            }
        }

        public int EditingControlRowIndex {
            get {
                return rowIndex;
            }
            set {
                rowIndex = value;
            }
        }

        public bool EditingControlValueChanged {
            get {
                return valueChanged;
            }
            set {
                valueChanged = value;
            }
        }

        public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey) {
            switch (key & Keys.KeyCode) {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                case Keys.Delete: {
                        DataGridViewSelectedRowCollection selectedRows = dataGridView.SelectedRows;
                        if (selectedRows.Count > 0) {
                            DataGridViewRow selectedRow = selectedRows[selectedRows.Count - 1];
                            dataGridView.Rows.RemoveAt(selectedRow.HeaderCell.RowIndex);
                        }
                        return true;
                    }
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        public Cursor EditingPanelCursor {
            get {
                return base.Cursor;
            }
        }

        public void PrepareEditingControlForEdit(bool selectAll) {
        }

        public bool RepositionEditingControlOnValueChange {
            get {
                return false;
            }
        }

        protected void ComboboxColor_SelectedIndexChanged(object sender, EventArgs e) {
            if (dataGridView.SelectedCells != null && dataGridView.SelectedCells.Count > 0)
                dataGridView.SelectedCells[0].Value = this.Text;
        }

        public static void DrawComboboxItemColour(ComboBox cbColour, Brush boxColour, String itemDescription, DrawItemEventArgs e) {
            try {
                e.Graphics.FillRectangle(new SolidBrush(cbColour.BackColor), e.Bounds);
                e.DrawBackground();
                Boolean comboEnabled = cbColour.Enabled;

                // Write colour name
                Boolean highlighted = (e.State & DrawItemState.Selected) != DrawItemState.None;
                Brush brush = comboEnabled ? SystemBrushes.WindowText : SystemBrushes.InactiveCaptionText;
                if (highlighted)
                    brush = comboEnabled ? SystemBrushes.HighlightText : SystemBrushes.InactiveCaptionText;

                DrawComboboxItemColour(comboEnabled, boxColour, brush, itemDescription, e.Graphics, e.Bounds);

                // Draw the focus rectangle if appropriate
                if ((e.State & DrawItemState.NoFocusRect) == DrawItemState.None)
                    e.DrawFocusRectangle();
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }

        public static void DrawComboboxItemColour(Boolean comboEnabled, Brush boxColour, Brush textColour, String itemDescription, Graphics graphics, Rectangle cellBounds) {
            try {
                // Draw colour box
                Rectangle colourbox = new Rectangle {
                    X = cellBounds.X + 2,
                    Y = cellBounds.Y + 2,
                    Height = cellBounds.Height - 5,
                    Width = 18
                };
                graphics.FillRectangle(boxColour, colourbox);
                graphics.DrawRectangle(comboEnabled ? SystemPens.WindowText : SystemPens.InactiveBorder, colourbox);

                int textX = cellBounds.X + colourbox.X + colourbox.Width + 2;

                graphics.DrawString(itemDescription, Control.DefaultFont, textColour,
                    /*cellBounds.X*/ +colourbox.X + colourbox.Width + 2,
                    cellBounds.Y + ((cellBounds.Height - Control.DefaultFont.Height) / 2));

            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }
    }

    public class OutlookColourCombobox : ColourCombobox, IDataGridViewEditingControl {
        public OutlookColourCombobox() {
            PopulateDropdownItems();

            this.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DrawItem += new DrawItemEventHandler(ComboboxColor_DrawItem);
            this.SelectedIndexChanged += new EventHandler(base.ComboboxColor_SelectedIndexChanged);
        }

        public void PopulateDropdownItems() {
            Dictionary<Outlook.Categories.ColourInfo, String> cbItems = new Dictionary<Outlook.Categories.ColourInfo, String>();
            foreach (Outlook.Categories.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                cbItems.Add(ci, ci.Text);
            }
            this.DataSource = new BindingSource(cbItems, null);
            this.DisplayMember = "Value";
            this.ValueMember = "Key";
        }

        void ComboboxColor_DrawItem(object sender, DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            int indexItem = e.Index;
            if (indexItem < 0 || indexItem >= cbColour.Items.Count)
                return;

            KeyValuePair<Outlook.Categories.ColourInfo, String> kvp = (KeyValuePair<Outlook.Categories.ColourInfo, String>)cbColour.Items[indexItem];
            if (kvp.Key != null) {
                // Get the colour
                OlCategoryColor olColour = kvp.Key.OutlookCategory;
                Brush brush = new SolidBrush(Outlook.Categories.Map.RgbColour(olColour));

                DrawComboboxItemColour(cbColour, brush, kvp.Value, e);
            }
        }
    }

    public class GoogleColourCombobox : ColourCombobox, IDataGridViewEditingControl {
        public GoogleColourCombobox() {
            PopulateDropdownItems();

            this.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DrawItem += new DrawItemEventHandler(ComboboxColor_DrawItem);
            this.SelectedIndexChanged += new EventHandler(base.ComboboxColor_SelectedIndexChanged);
        }

        public void PopulateDropdownItems() {
            Dictionary<Ogcs.Google.EventColour.Palette, String> cbItems = new Dictionary<Ogcs.Google.EventColour.Palette, String>();
            foreach (Ogcs.Google.EventColour.Palette ci in Forms.ColourMap.GoogleComboBox.Items) {
                cbItems.Add(ci, ci.Name);
            }
            this.DataSource = new BindingSource(cbItems, null);
            this.DisplayMember = "Value";
            this.ValueMember = "Key";
        }

        void ComboboxColor_DrawItem(object sender, DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            if (e.Index < 0 || e.Index >= cbColour.Items.Count)
                return;

            KeyValuePair<Ogcs.Google.EventColour.Palette, String> kvp = (KeyValuePair<Ogcs.Google.EventColour.Palette, String>)cbColour.Items[e.Index];
            if (kvp.Key != null) {
                // Get the colour
                Brush brush = new SolidBrush(kvp.Key.RgbValue);

                DrawComboboxItemColour(cbColour, brush, kvp.Value, e);
            }
        }
    }
    #endregion
}
