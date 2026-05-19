/// FILE: Safety_System/CoreTable/App_CoreTable.UI.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_CoreTable
    {
        private Control BuildUI()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            Padding lblPad = new Padding(0, 8, 5, 0); 
            Padding ctrlPad = new Padding(0, 4, 5, 0); 
            Padding btnPad = new Padding(0, 0, 10, 0); 
            int btnHeight = 35; 

            GroupBox boxTop = new GroupBox { 
                Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                AutoSize = true, 
                AutoSizeMode = AutoSizeMode.GrowAndShrink, 
                Padding = new Padding(10, 15, 10, 10), 
                Margin = new Padding(0, 0, 0, 10) 
            };
            
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };

            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = lblPad };
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 25; i <= currentYear + 25; i++) {
                _cboStartYear.Items.Add(i); _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2")); _cboEndMonth.Items.Add(i.ToString("D2"));
            }
            for (int i = 1; i <= 31; i++) {
                _cboStartDay.Items.Add(i.ToString("D2")); _cboEndDay.Items.Add(i.ToString("D2"));
            }

            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, (_timeMode == TimeMode.Date) ? DateTime.Today.AddDays(-30) : DateTime.Today.AddMonths(-6));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            _btnRead = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            _btnRead.FlatAppearance.BorderSize = 0;
            
            Label lblLatestCount = new Label { Text = "最近筆數:", AutoSize = true, Margin = new Padding(15, 8, 5, 0) };
            _txtLatestCount = new TextBox { Width = 50, Text = "50", Margin = ctrlPad }; 
            
            Button bLimitRead = new Button { Text = "查詢", Size = new Size(80, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bLimitRead.FlatAppearance.BorderSize = 0;
            
            Label lblSY = new Label { Text = "年", AutoSize = true, Margin = lblPad };
            Label lblSM = new Label { Text = "月", AutoSize = true, Margin = lblPad };
            Label lblSD = new Label { Text = "日", AutoSize = true, Margin = lblPad };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            Label lblEY = new Label { Text = "年", AutoSize = true, Margin = lblPad };
            Label lblEM = new Label { Text = "月", AutoSize = true, Margin = lblPad };
            Label lblED = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 15, 0) }; 

            if (_timeMode == TimeMode.YearMonth || _timeMode == TimeMode.Year) {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
            }
            if (_timeMode == TimeMode.Year) {
                _cboStartMonth.Visible = false; lblSM.Visible = false;
                _cboEndMonth.Visible = false; lblEM.Visible = false;
            }

            _btnToggle = new Button { Text = "+", Size = new Size(45, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(142, 142, 147), ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold) };
            _btnToggle.FlatAppearance.BorderSize = 0;
            _btnToggle.Click += delegate(object s, EventArgs e) { 
                _boxAdvanced.Visible = !_boxAdvanced.Visible; 
                _btnToggle.Text = _boxAdvanced.Visible ? "-" : "+"; 
            };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存", Size = new Size(95, btnHeight), Margin = new Padding(0, 0, 10, 0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.FlatAppearance.BorderSize = 0;

            flpTop.Controls.AddRange(new Control[] {
                lblRange, _cboStartYear, lblSY, _cboStartMonth, lblSM, _cboStartDay, lblSD,
                lblTilde, _cboEndYear, lblEY, _cboEndMonth, lblEM, _cboEndDay, lblED,
                _btnRead, lblLatestCount, _txtLatestCount, bLimitRead,
                _btnToggle, _btnSave
            });
            boxTop.Controls.Add(flpTop);

            _boxAdvanced = new GroupBox { 
                Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), 
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Visible = false, 
                Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) 
            };
            
            TableLayoutPanel tlpAdvLeft = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            tlpAdvLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpAdvLeft.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            tlpAdvLeft.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            FlowLayoutPanel flpAdvRow1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            FlowLayoutPanel flpAdvRow2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 0) };

            _txtNewColName = new TextBox { Width = 110, Margin = ctrlPad };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(95, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bAdd.FlatAppearance.BorderSize = 0;
            bAdd.Click += delegate(object s, EventArgs e) { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { 
                    UnfreezeAllColumns();
                    DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); 
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (!dt.Columns.Contains(_txtNewColName.Text)) dt.Columns.Add(_txtNewColName.Text, typeof(string));
                    
                    _isApplyingWidths = true;
                    _dgv.SuspendLayout();
                    ApplyGridStyles(); 
                    UpdateCboColumns(); 
                    _dgv.ResumeLayout(true);
                    _isApplyingWidths = false;

                    _txtNewColName.Clear(); 
                    ApplyFreezeState();
                    MessageBox.Show("欄位新增成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } 
            };
            
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad }; 
            _txtRenameCol = new TextBox { Width = 120, Margin = ctrlPad };
            Button bRen = new Button { Text = "標題更改", Size = new Size(95, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bRen.FlatAppearance.BorderSize = 0;
            bRen.Click += delegate(object s, EventArgs e) { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { 
                    UnfreezeAllColumns();
                    string oldName = _cboColumns.SelectedItem.ToString();
                    DataManager.RenameColumn(_dbName, _tableName, oldName, _txtRenameCol.Text); 
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (dt.Columns.Contains(oldName)) dt.Columns[oldName].ColumnName = _txtRenameCol.Text;
                    if (_dgv.Columns.Contains(oldName)) { _dgv.Columns[oldName].HeaderText = _txtRenameCol.Text; _dgv.Columns[oldName].Name = _txtRenameCol.Text; }
                    UpdateCboColumns(); 
                    _txtRenameCol.Clear(); 
                    ApplyFreezeState();
                    MessageBox.Show("欄位名稱修改成功！");
                } 
            };
            
            Button bDelCol = new Button { Text = "刪除欄", Size = new Size(80, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 59, 48), ForeColor = Color.White };
            bDelCol.FlatAppearance.BorderSize = 0;
            bDelCol.Click += delegate(object s, EventArgs e) { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { 
                    string colToDrop = _cboColumns.SelectedItem.ToString();
                    if(MessageBox.Show($"確定刪除整欄【{colToDrop}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) { 
                        UnfreezeAllColumns();
                        DataManager.DropColumn(_dbName, _tableName, colToDrop); 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (dt.Columns.Contains(colToDrop)) dt.Columns.Remove(colToDrop);
                        if (_dgv.Columns.Contains(colToDrop)) _dgv.Columns.Remove(colToDrop);
                        UpdateCboColumns(); 
                        ApplyFreezeState();
                        MessageBox.Show("欄位刪除成功！");
                    } 
                } 
            };
            
            Button bDelRow = new Button { Text = "🗑 刪除列", Size = new Size(110, btnHeight), Margin = new Padding(0,0,0,0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 59, 48), ForeColor = Color.White };
            bDelRow.FlatAppearance.BorderSize = 0;
            // 事件將於 App_CoreTable.Events 綁定

            _btnColSettings = new Button { Text = "👁️ 顯示設定", Size = new Size(120, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(142, 142, 147), ForeColor = Color.White };
            _btnColSettings.FlatAppearance.BorderSize = 0;

            flpAdvRow1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = lblPad }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            flpAdvRow1.Controls.Add(new Panel { Width = 30, Height = 1 }); 
            flpAdvRow1.Controls.Add(_btnColSettings);

            _cboSearchColumn = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _txtSearchKeyword = new TextBox { Width = 180, Margin = ctrlPad };
            _btnAdvancedSearch = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = new Padding(0,0,0,0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            _btnAdvancedSearch.FlatAppearance.BorderSize = 0;

            _btnImport = new Button { Text = "📥 匯入", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 149, 0), ForeColor = Color.White };
            _btnImport.FlatAppearance.BorderSize = 0;

            _btnExport = new Button { Text = "📤 匯出", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
            _btnExport.FlatAppearance.BorderSize = 0;

            _btnExportPdf = new Button { Text = "📄 導出 PDF", Size = new Size(120, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
            _btnExportPdf.FlatAppearance.BorderSize = 0;

            if (_logic is LawLogic) {
                _btnRtfToExcel = new Button { Text = "📄 RTF轉 Excel", Size = new Size(160, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
                _btnRtfToExcel.FlatAppearance.BorderSize = 0;
                flpAdvRow2.Controls.Add(_btnRtfToExcel);
            }

            flpAdvRow2.Controls.AddRange(new Control[] { new Label { Text = "查詢資料:", AutoSize = true, Margin = lblPad }, _cboSearchColumn, new Label { Text = "關鍵字(含):", AutoSize = true, Margin = lblPad }, _txtSearchKeyword, _btnAdvancedSearch });
            flpAdvRow2.Controls.Add(new Panel { Width = 30, Height = 1 }); 
            flpAdvRow2.Controls.AddRange(new Control[] { _btnImport, _btnExport, _btnExportPdf });

            tlpAdvLeft.Controls.Add(flpAdvRow1, 0, 0);
            tlpAdvLeft.Controls.Add(flpAdvRow2, 0, 1);
            _boxAdvanced.Controls.Add(tlpAdvLeft);

            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AllowUserToResizeColumns = true, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells, AllowUserToOrderColumns = true, Margin = new Padding(0, 10, 0, 10),
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };
            _dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            InitContextMenu(bDelRow); 

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            // 將按鈕暫存至動態變數供 BindEvents 綁定
            _btnRead.Tag = bLimitRead; 

            return main;
        }

        private void InitContextMenu(Button btnDelRow)
        {
            _ctxMenu = new ContextMenuStrip { Font = new Font("Microsoft JhengHei UI", 11F) };

            ToolStripMenuItem itemSave = new ToolStripMenuItem("💾 儲存");
            ToolStripMenuItem itemDeleteRow = new ToolStripMenuItem("🗑️ 刪除選取列");
            ToolStripMenuItem itemColSettings = new ToolStripMenuItem("👁️ 顯示設定");
            ToolStripMenuItem itemFreeze = new ToolStripMenuItem("❄️ 凍結此欄(含)以左視窗");
            ToolStripMenuItem itemUnfreeze = new ToolStripMenuItem("🔥 取消凍結");
            ToolStripMenuItem itemImport = new ToolStripMenuItem("📥 匯入");
            ToolStripMenuItem itemExport = new ToolStripMenuItem("📤 匯出");
            ToolStripMenuItem itemPdf = new ToolStripMenuItem("📄 導出 PDF");

            itemSave.Tag = "Save";
            itemDeleteRow.Tag = "DeleteRow";
            itemColSettings.Tag = "ColSettings";
            itemImport.Tag = "Import";
            itemExport.Tag = "Export";
            itemPdf.Tag = "Pdf";

            itemFreeze.Click += delegate(object s, EventArgs e) {
                if (_rightClickedColIndex >= 0 && _rightClickedColIndex < _dgv.Columns.Count) {
                    _frozenColumnName = _dgv.Columns[_rightClickedColIndex].Name;
                    ApplyFreezeState();
                }
            };

            itemUnfreeze.Click += delegate(object s, EventArgs e) {
                _frozenColumnName = null;
                UnfreezeAllColumns();
            };

            _ctxMenu.Items.AddRange(new ToolStripItem[] { 
                itemSave, itemDeleteRow, new ToolStripSeparator(),
                itemFreeze, itemUnfreeze, new ToolStripSeparator(),
                itemColSettings, new ToolStripSeparator(),
                itemImport, itemExport, itemPdf 
            });

            _ctxMenu.Tag = btnDelRow;
        }

        private void SetUIState(bool isEnabled, string statusText, Color statusColor) 
        {
            _btnRead.Enabled = isEnabled; 
            _btnSave.Enabled = isEnabled; 
            _btnImport.Enabled = isEnabled; 
            _btnExport.Enabled = isEnabled; 
            _btnExportPdf.Enabled = isEnabled; 
            _btnColSettings.Enabled = isEnabled; 
            _btnAdvancedSearch.Enabled = isEnabled; 
            if(_btnRtfToExcel != null) _btnRtfToExcel.Enabled = isEnabled;
            
            _lblStatus.Text = statusText; 
            _lblStatus.ForeColor = statusColor;
        }

        private void ApplyGridStyles() 
        {
            if (_dgv.Columns.Contains("Id")) {
                _dgv.Columns["Id"].ReadOnly = true;
                _dgv.Columns["Id"].Visible = false;
            }
            
            if (_dgv.Columns.Contains(_dateColumnName)) {
                string fmt = "yyyy-MM-dd";
                if (_timeMode == TimeMode.YearMonth) fmt = "yyyy-MM";
                else if (_timeMode == TimeMode.Year) fmt = "yyyy";
                _dgv.Columns[_dateColumnName].DefaultCellStyle.Format = fmt;
            }
            
            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (_columnVisibility.ContainsKey(col.Name)) col.Visible = _columnVisibility[col.Name];

                if (col.Name.Contains("附件檔案")) {
                    col.ReadOnly = true; 
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                } else {
                    col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }

            SetupDropdownColumns();
            
            _dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.DisplayedCells);
            _dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            foreach (DataGridViewColumn col in _dgv.Columns) {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                if (_columnWidths.ContainsKey(col.Name) && _columnWidths[col.Name] > 0) {
                    col.Width = _columnWidths[col.Name];
                } else if (_logic is LawLogic) {
                    if (col.Name == "法規名稱") col.Width = 250;
                    else if (col.Name == "內容") col.Width = 400;
                    else if (col.Name == "重點摘要") col.Width = 200;
                } else {
                    if (col.Width < 80) col.Width = 80;
                }
            }
        }

        private void SetupDropdownColumns() 
        {
            List<DataGridViewColumn> cols = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn col in _dgv.Columns) cols.Add(col);

            foreach (DataGridViewColumn col in cols) {
                string[] items = _logic.GetDropdownList(_tableName, col.Name);
                string[] dbItems = App_DropdownManager.GetAllOptionsForColumn(_tableName, col.Name);
                
                if (dbItems != null && dbItems.Length > 1) { items = dbItems; }

                if (items != null && items.Length > 0 && !(_dgv.Columns[col.Name] is DataGridViewComboBoxColumn)) {
                    int colIndex = col.Index; 
                    _dgv.Columns.RemoveAt(colIndex);
                    
                    DataGridViewComboBoxColumn cboCol = new DataGridViewComboBoxColumn { 
                        Name = col.Name, HeaderText = col.HeaderText, DataPropertyName = col.DataPropertyName, 
                        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox, FlatStyle = FlatStyle.Flat, 
                        SortMode = DataGridViewColumnSortMode.Automatic 
                    };
                    
                    if (_columnVisibility.ContainsKey(col.Name)) cboCol.Visible = _columnVisibility[col.Name];
                    
                    List<string> finalItems = new List<string>();
                    foreach (string item in items) finalItems.Add(item);

                    if (_dgv.DataSource is DataTable) { 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        foreach (DataRow row in dt.Rows) { 
                            object valObj = row[col.Name];
                            if (valObj != null) {
                                string val = valObj.ToString().Trim(); 
                                if (!string.IsNullOrEmpty(val) && !finalItems.Contains(val)) finalItems.Add(val); 
                            }
                        } 
                    }
                    
                    cboCol.Items.AddRange(finalItems.ToArray()); 
                    _dgv.Columns.Insert(colIndex, cboCol);
                }
            }
        }

        private void PreFillComboBoxItems(DataTable dt)
        {
            if (dt == null || _dgv.Columns.Count == 0) return;
            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col is DataGridViewComboBoxColumn && dt.Columns.Contains(col.Name)) {
                    DataGridViewComboBoxColumn cboCol = (DataGridViewComboBoxColumn)col;
                    List<string> existingItems = new List<string>();
                    foreach (object item in cboCol.Items) existingItems.Add(item.ToString());

                    foreach (DataRow row in dt.Rows) {
                        if (row.RowState == DataRowState.Deleted) continue;
                        object valObj = row[col.Name];
                        if (valObj != null) {
                            string val = valObj.ToString().Trim();
                            if (!string.IsNullOrEmpty(val) && !existingItems.Contains(val)) {
                                existingItems.Add(val); 
                                cboCol.Items.Add(val);
                            }
                        }
                    }
                }
            }
        }

        private void UpdateCboColumns() 
        {
            string currentSearchSel = "";
            if (_cboSearchColumn.SelectedItem != null) {
                currentSearchSel = _cboSearchColumn.SelectedItem.ToString();
            }
            _cboColumns.Items.Clear(); 
            _cboSearchColumn.Items.Clear(); 
            _cboSearchColumn.Items.Add(""); 
            
            foreach (DataGridViewColumn c in _dgv.Columns) { 
                if (c.Name != "Id" && c.Name != _dateColumnName) { _cboColumns.Items.Add(c.Name); } 
                if (c.Name != "Id") { _cboSearchColumn.Items.Add(c.Name); } 
            }
            
            if (!string.IsNullOrEmpty(currentSearchSel) && _cboSearchColumn.Items.Contains(currentSearchSel)) { 
                _cboSearchColumn.SelectedItem = currentSearchSel; 
            } else if (_cboSearchColumn.Items.Count > 0) { 
                _cboSearchColumn.SelectedIndex = 0; 
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) 
        {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2"); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private class ColDisplayIndexComparerDesc : IComparer<DataGridViewColumn> 
        {
            public int Compare(DataGridViewColumn x, DataGridViewColumn y) {
                return y.DisplayIndex.CompareTo(x.DisplayIndex);
            }
        }

        private class ColDisplayIndexComparerAsc : IComparer<DataGridViewColumn> 
        {
            public int Compare(DataGridViewColumn x, DataGridViewColumn y) {
                return x.DisplayIndex.CompareTo(y.DisplayIndex);
            }
        }

        private void UnfreezeAllColumns()
        {
            if (_dgv == null || _dgv.Columns.Count == 0) return;

            List<DataGridViewColumn> cols = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn c in _dgv.Columns) {
                cols.Add(c);
            }
            
            cols.Sort(new ColDisplayIndexComparerDesc());

            foreach (DataGridViewColumn col in cols) {
                col.Frozen = false;
            }
        }

        private void ApplyFreezeState()
        {
            if (string.IsNullOrEmpty(_frozenColumnName) || _dgv == null || !_dgv.Columns.Contains(_frozenColumnName)) return;

            try {
                UnfreezeAllColumns(); 
                int targetIndex = _dgv.Columns[_frozenColumnName].DisplayIndex;
                
                List<DataGridViewColumn> colsToFreeze = new List<DataGridViewColumn>();
                foreach (DataGridViewColumn c in _dgv.Columns) {
                    if (c.DisplayIndex <= targetIndex) {
                        colsToFreeze.Add(c);
                    }
                }
                
                colsToFreeze.Sort(new ColDisplayIndexComparerAsc());
                                      
                foreach(DataGridViewColumn col in colsToFreeze) {
                    col.Frozen = true;
                }
            } 
            catch { }
        }
    }
}
