using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq; // 🟢 支援排序操作
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_WaterChemicals
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        private const string DbName = "Water"; 
        private const string TableName = "WaterChemicals"; 

        public Control GetView()
        {
            // 🟢 更新為新的資料表結構
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WaterChemicals] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [PAC] TEXT, 
                [NAOH] TEXT, 
                [高分子] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // 🟢 更新標題
            GroupBox boxTop = new GroupBox { Text = $"水處理用藥記錄 (庫：{DbName} 表：{TableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-30) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
            
            Button bRead = new Button { Text = "讀取資料", Size = new Size(120, 35) };
            bRead.Click += (s, e) => { 
                RefreshGrid(); 
                if (!_isFirstLoad) { 
                    int count = ((DataTable)_dgv.DataSource).Rows.Count; 
                    MessageBox.Show(Form.ActiveForm, $"讀取完成！共找到 {count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                } 
            };

            Button bSave = new Button { 
                Name = "btnSave", 
                Text = "💾 儲存", 
                Size = new Size(120, 35), 
                BackColor = Color.ForestGreen, 
                ForeColor = Color.White, 
                Margin = new Padding(30, 0, 0, 0) 
            };
            bSave.Click += (s, e) => {
                _dgv.EndEdit(); 
                
                // 🟢 1. 無條件先存排序 (避免因為沒修改資料而略過)
                SaveColumnOrder(); 

                // 🟢 2. 儲存資料
                if (DataManager.ValidateAndSaveTable(DbName, TableName, (DataTable)_dgv.DataSource)) {
                    MessageBox.Show(Form.ActiveForm, "儲存完成！(已記憶最新欄位排序)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshGrid();
                } else {
                    // 若資料沒有異動，也要明確告訴使用者「版面有存起來」
                    MessageBox.Show(Form.ActiveForm, "欄位排序已儲存！(資料無異動)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            
            Button bExport = new Button { Text = "匯出", Size = new Size(120, 35) };
            bExport.Click += BtnExport_Click;

            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(120, 35) };
            bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
                _btnToggle.BackColor = _boxAdvanced.Visible ? Color.LightCoral : Color.LightGray;
            };

            row1.Controls.AddRange(new Control[] { lblRange, _dtpStart, lblTilde, _dtpEnd, bRead, bExport, bImport, _btnToggle, bSave });
            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdvMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            FlowLayoutPanel row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Label lblOps = new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; 
            _txtNewColName = new TextBox { Width = 120 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) };
            bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && VerifyPassword()) { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };

            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "標題更改", Size = new Size(120, 35) };
            bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && VerifyPassword()) { DataManager.RenameColumn(DbName, TableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };

            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null) {
                    string colToDrop = _cboColumns.SelectedItem.ToString();
                    if (MessageBox.Show(Form.ActiveForm, $"警告：確定要刪除整欄【{colToDrop}】嗎？\n刪除後該欄位所有歷史資料將永久遺失！", "刪除欄位確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        if (VerifyPassword()) { DataManager.DropColumn(DbName, TableName, colToDrop); RefreshGrid(); } 
                    }
                } 
            };

            Button bDelRow = new Button { Text = "刪除整列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += (s, e) => { if (_dgv.CurrentRow != null && _dgv.CurrentRow.Cells["Id"].Value != DBNull.Value && VerifyPassword()) { DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value)); RefreshGrid(); } };

            row2.Controls.AddRange(new Control[] { lblOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });

            FlowLayoutPanel row3 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            Label lblLatest = new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            TextBox txtLatestCount = new TextBox { Width = 120, Text = "50" }; 
            
            Button bReadLatest = new Button { Text = "讀取筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bReadLatest.Click += (s, e) => {
                if (int.TryParse(txtLatestCount.Text, out int limit) && limit > 0) {
                    DataTable dt = DataManager.GetLatestRecords(DbName, TableName, limit);
                    _dgv.DataSource = dt;
                    
                    if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
                    _cboColumns.Items.Clear();
                    foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);

                    RestoreColumnOrder(); // 讀取自訂筆數後，立刻套用排序

                    MessageBox.Show(Form.ActiveForm, $"讀取完成！共載入最近 {dt.Rows.Count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    if (dt.Rows.Count > 0) {
                        DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                        foreach(DataRow r in dt.Rows) {
                            if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { 
                                if (d < minD) minD = d; 
                                if (d > maxD) maxD = d; 
                            }
                        }
                        if (minD <= maxD) { _dtpStart.Value = minD; _dtpEnd.Value = maxD; }
                    }
                } else {
                    MessageBox.Show(Form.ActiveForm, "請輸入有效的正整數！", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            row3.Controls.AddRange(new Control[] { lblLatest, txtLatestCount, bReadLatest });

            flpAdvMain.Controls.Add(row2);
            flpAdvMain.Controls.Add(row3);
            _boxAdvanced.Controls.Add(flpAdvMain);

            // 🟢 開啟 AllowUserToOrderColumns 允許使用者拖曳移動欄位
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToOrderColumns = true };
            
            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(_boxAdvanced, 0, 1);
            main.Controls.Add(_dgv, 0, 2);

            RefreshGrid();
            return main;
        }

        private void RefreshGrid() {
            if (_isFirstLoad) {
                DataTable dt = DataManager.GetLatestRecords(DbName, TableName, 30);
                _dgv.DataSource = dt;
                if (dt.Rows.Count == 0) MessageBox.Show(Form.ActiveForm, "【系統連線成功】目前資料表尚無任何紀錄。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else {
                    DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                    foreach(DataRow r in dt.Rows) if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { if (d < minD) minD = d; if (d > maxD) maxD = d; }
                    if (minD <= maxD) { _dtpStart.Value = minD; _dtpEnd.Value = maxD; }
                }
                _isFirstLoad = false;
            } else {
                _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            }
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);

            // 🟢 每次刷新表格後，自動套用記憶的欄位排序
            RestoreColumnOrder();
        }

        // ==========================================
        // 🟢 欄位排序記憶功能：儲存
        // ==========================================
        private void SaveColumnOrder() {
            try {
                var orderedCols = _dgv.Columns.Cast<DataGridViewColumn>()
                                      .OrderBy(c => c.DisplayIndex)
                                      .Select(c => c.Name).ToArray();
                
                string fileName = $"ColOrder_{DbName}_{TableName}.txt";
                File.WriteAllText(fileName, string.Join(",", orderedCols), Encoding.UTF8);

                if (_dgv.DataSource is DataTable dt) {
                    for (int i = 0; i < orderedCols.Length; i++) {
                        if (dt.Columns.Contains(orderedCols[i])) {
                            dt.Columns[orderedCols[i]].SetOrdinal(i);
                        }
                    }
                }
            } catch { }
        }

        // ==========================================
        // 🟢 欄位排序記憶功能：讀取
        // ==========================================
        private void RestoreColumnOrder() {
            try {
                string fileName = $"ColOrder_{DbName}_{TableName}.txt";
                if (File.Exists(fileName)) {
                    string[] savedCols = File.ReadAllText(fileName, Encoding.UTF8).Split(',');
                    for (int i = 0; i < savedCols.Length; i++) {
                        if (_dgv.Columns.Contains(savedCols[i])) {
                            _dgv.Columns[savedCols[i]].DisplayIndex = i;
                        }
                    }

                    if (_dgv.DataSource is DataTable dt) {
                        var orderedCols = _dgv.Columns.Cast<DataGridViewColumn>()
                                              .OrderBy(c => c.DisplayIndex)
                                              .Select(c => c.Name).ToArray();
                        for (int i = 0; i < orderedCols.Length; i++) {
                            if (dt.Columns.Contains(orderedCols[i])) {
                                dt.Columns[orderedCols[i]].SetOrdinal(i);
                            }
                        }
                    }
                }
            } catch { }
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.Add(lbl); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            
            return p.ShowDialog(Form.ActiveForm) == DialogResult.OK && t.Text == "tces";
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            if (_dgv.Rows.Count == 0 || (_dgv.Rows.Count == 1 && _dgv.Rows[0].IsNewRow)) { MessageBox.Show(Form.ActiveForm, "沒有資料可匯出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            
            if (_dgv.DataSource is DataTable dTable) {
                var orderedCols = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray();
                for (int i = 0; i < orderedCols.Length; i++) {
                    if (dTable.Columns.Contains(orderedCols[i])) dTable.Columns[orderedCols[i]].SetOrdinal(i);
                }
            }

            // 🟢 更新匯出檔案名稱前綴
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|CSV 檔案 (*.csv)|*.csv", FileName = "水處理用藥記錄_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog(Form.ActiveForm) == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) { using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName)); } }
                        else {
                            StringBuilder sb = new StringBuilder();
                            string[] colNames = new string[dt.Columns.Count];
                            for (int i = 0; i < dt.Columns.Count; i++) colNames[i] = dt.Columns[i].ColumnName;
                            sb.AppendLine(string.Join(",", colNames));
                            foreach (DataRow row in dt.Rows) { string[] fields = new string[dt.Columns.Count]; for (int i = 0; i < dt.Columns.Count; i++) fields[i] = row[i]?.ToString().Replace(",", "，"); sb.AppendLine(string.Join(",", fields)); }
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show(Form.ActiveForm, "資料匯出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show(Form.ActiveForm, "匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案 (*.csv)|*.csv" }) {
                if (ofd.ShowDialog(Form.ActiveForm) == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length < 2) return; 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) { string cn = headers[h].Trim(); if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim(); }
                            dt.Rows.Add(nr);
                        }
                        MessageBox.Show(Form.ActiveForm, $"載入 {lines.Length - 1} 筆！請記得按儲存。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show(Form.ActiveForm, "匯入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}
