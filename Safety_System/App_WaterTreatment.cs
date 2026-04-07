/// FILE: Safety_System/App_WaterTreatment.cs ///
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq; 
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_WaterTreatment
    {
        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        private const string DbName = "Water"; 
        private const string TableName = "WaterMeterReadings"; 

        // 🟢 自動運算共用 Helper，負責星期轉換、相減統計及防呆控制
        private DataGridViewAutoCalcHelper _calcHelper; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WaterMeterReadings] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [星期] TEXT,
                [用電量] TEXT, 
                [用電量日統計] TEXT, 
                [廢水進流量] TEXT, 
                [廢水處理量日統計] TEXT, 
                [廢水處理量] TEXT, 
                [水站廢水排放量] TEXT, 
                [水站廢水排放量日統計] TEXT, 
                [納管排放量] TEXT,
                [納管排放量日統計] TEXT,
                [納廢回收6吋] TEXT, 
                [納廢回收6吋日統計] TEXT, 
                [雙介質A] TEXT,
                [雙介質A日統計] TEXT, 
                [雙介質B] TEXT, 
                [雙介質B日統計] TEXT, 
                [雙介質AB日統計] TEXT, 
                [軟水A] TEXT, 
                [軟水A日統計] TEXT, 
                [軟水B] TEXT, 
                [軟水B日統計] TEXT, 
                [軟水C] TEXT,
                [軟水C日統計] TEXT, 
                [濃縮水至冷卻水池] TEXT,
                [濃縮水至冷卻水池日統計] TEXT,
                [濃縮水至逆洗池] TEXT,
                [濃縮水至逆洗池日統計] TEXT,
                [廠區自來水] TEXT,
                [廠區自來水日統計] TEXT,
                [污泥產出包數] TEXT,
                [備註] TEXT);");

            // 系統啟動時若發現沒有星期欄位，自動補上
            if (!DataManager.GetColumnNames(DbName, TableName).Contains("星期"))
            {
                DataManager.AddColumn(DbName, TableName, "星期");
            }

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"廢水處理水量記錄 (庫：{DbName} 表：{TableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 25; i <= currentYear + 25; i++) {
                _cboStartYear.Items.Add(i);
                _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2"));
                _cboEndMonth.Items.Add(i.ToString("D2"));
            }
            for (int i = 1; i <= 31; i++) {
                _cboStartDay.Items.Add(i.ToString("D2"));
                _cboEndDay.Items.Add(i.ToString("D2"));
            }

            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            Label lblStartYear = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblStartMonth = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblStartDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEndYear = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEndMonth = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEndDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

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
                Margin = new Padding(0, 3, 3, 3) 
            };
            bSave.Click += (s, e) => {
                _dgv.EndEdit(); 
                SaveColumnOrder(); 

                DataTable dtToSave = (DataTable)_dgv.DataSource;
                EnforceMonthFormat(dtToSave);

                if (DataManager.ValidateAndSaveTable(DbName, TableName, dtToSave)) {
                    MessageBox.Show(Form.ActiveForm, "儲存完成！(已記憶最新欄位排序)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshGrid();
                } else {
                    MessageBox.Show(Form.ActiveForm, "欄位排序已儲存！(資料無異動)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            
            Button bExport = new Button { Text = "匯出", Size = new Size(100, 35) };
            bExport.Click += BtnExport_Click;

            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(120, 35) };
            bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
                _btnToggle.BackColor = _boxAdvanced.Visible ? Color.LightCoral : Color.LightGray;
            };

            row1.Controls.AddRange(new Control[] { 
                lblRange, 
                _cboStartYear, lblStartYear, _cboStartMonth, lblStartMonth, _cboStartDay, lblStartDay, 
                lblTilde, 
                _cboEndYear, lblEndYear, _cboEndMonth, lblEndMonth, _cboEndDay, lblEndDay, 
                bRead, bExport, bImport, _btnToggle, bSave 
            });
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
                    EnforceMonthFormat(dt);
                    _dgv.DataSource = dt;
                    
                    if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
                    _cboColumns.Items.Clear();
                    foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);

                    RestoreColumnOrder(); 

                    MessageBox.Show(Form.ActiveForm, $"讀取完成！共載入最近 {dt.Rows.Count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    if (dt.Rows.Count > 0) {
                        DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                        foreach(DataRow r in dt.Rows) {
                            if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { 
                                if (d < minD) minD = d; 
                                if (d > maxD) maxD = d; 
                            }
                        }
                        if (minD <= maxD) { 
                            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, minD); 
                            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, maxD); 
                        }
                    }
                } else {
                    MessageBox.Show(Form.ActiveForm, "請輸入有效的正整數！", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            row3.Controls.AddRange(new Control[] { lblLatest, txtLatestCount, bReadLatest });

            flpAdvMain.Controls.Add(row2);
            flpAdvMain.Controls.Add(row3);
            _boxAdvanced.Controls.Add(flpAdvMain);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToOrderColumns = true };
            
            // 🟢 加入 Ctrl+V 貼上事件攔截
            _dgv.KeyDown += Dgv_KeyDown;

            // 🟢 實例化全域運算共用 Helper 並綁定 DataGridView
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(_boxAdvanced, 0, 1);
            main.Controls.Add(_dgv, 0, 2);

            RefreshGrid();
            return main;
        }

        private void RefreshGrid() {
            if (_isFirstLoad) {
                DataTable dt = DataManager.GetLatestRecords(DbName, TableName, 30);
                EnforceMonthFormat(dt);
                _dgv.DataSource = dt;
                
                if (dt.Rows.Count == 0) MessageBox.Show(Form.ActiveForm, "【系統連線成功】目前資料表尚無任何紀錄。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else {
                    DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                    foreach(DataRow r in dt.Rows) if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { if (d < minD) minD = d; if (d > maxD) maxD = d; }
                    if (minD <= maxD) { 
                        SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, minD); 
                        SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, maxD); 
                    }
                }
                _isFirstLoad = false;
            } else {
                string sDate = GetStartDate().ToString("yyyy-MM-dd");
                string eDate = GetEndDate().ToString("yyyy-MM-dd");
                DataTable dt = DataManager.GetTableData(DbName, TableName, "日期", sDate, eDate);
                EnforceMonthFormat(dt);
                _dgv.DataSource = dt;
            }
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);

            RestoreColumnOrder();
        }

        private void EnforceMonthFormat(DataTable dt) {
            if (dt != null && dt.Columns.Contains("月份")) {
                foreach (DataRow row in dt.Rows) {
                    if (row.RowState == DataRowState.Deleted) continue;
                    string val = row["月份"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(val)) {
                        if (DateTime.TryParse(val, out DateTime parsedDate)) {
                            row["月份"] = parsedDate.ToString("yyyy-MM");
                        } else {
                            string normalized = val.Replace("/", "-");
                            var parts = normalized.Split('-');
                            if (parts.Length >= 2) {
                                if (int.TryParse(parts[0], out int y) && int.TryParse(parts[1], out int m)) {
                                    if (y < 100) y += 2000; 
                                    row["月份"] = $"{y:D4}-{m:D2}";
                                }
                            }
                        }
                    }
                }
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetStartDate() {
            return ParseComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
        }

        private DateTime GetEndDate() {
            return ParseComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);
        }

        private DateTime ParseComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime defaultDate) {
            if (y.SelectedItem == null || m.SelectedItem == null || d.SelectedItem == null) return defaultDate;
            if (int.TryParse(y.SelectedItem.ToString(), out int year) &&
                int.TryParse(m.SelectedItem.ToString(), out int month) &&
                int.TryParse(d.SelectedItem.ToString(), out int day)) {
                try {
                    int daysInMonth = DateTime.DaysInMonth(year, month);
                    if (day > daysInMonth) day = daysInMonth;
                    return new DateTime(year, month, day);
                } catch { return defaultDate; }
            }
            return defaultDate;
        }

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

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|CSV 檔案 (*.csv)|*.csv", FileName = "廢水處理水量記錄_" + DateTime.Now.ToString("yyyyMMdd") }) {
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

        // 🟢 加入改良版的 CSV 匯入，完美處理雙引號與千分位逗號，並套用防止當機控制
        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案 (*.csv)|*.csv" }) {
                if (ofd.ShowDialog(Form.ActiveForm) == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length < 2) return; 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        
                        string[] headers = ParseCsvLine(lines[0]);

                        _calcHelper?.BeginBulkUpdate(); // 開始匯入，關閉即時計算防當機

                        for (int i = 1; i < lines.Length; i++) {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            DataRow nr = dt.NewRow(); 
                            string[] vs = ParseCsvLine(lines[i]); // 使用改良版解析器
                            
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) { 
                                string cn = headers[h].Trim(); 
                                if (dt.Columns.Contains(cn) && cn != "Id") {
                                    // 確保雙引號被去除
                                    nr[cn] = vs[h].Trim().Trim('"'); 
                                }
                            }
                            dt.Rows.Add(nr);
                        }

                        _calcHelper?.EndBulkUpdate(); // 結束匯入，統一計算

                        MessageBox.Show(Form.ActiveForm, $"載入 {lines.Length - 1} 筆！請記得確認無誤後點擊「儲存」。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { 
                        _calcHelper?.EndBulkUpdate(); 
                        MessageBox.Show(Form.ActiveForm, "匯入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }

        // 🟢 新增：標準的 CSV 逐行解析器 (能正確處理雙引號內的千分位逗號)
        private string[] ParseCsvLine(string line)
        {
            var result = new System.Collections.Generic.List<string>();
            bool inQuotes = false;
            var currentField = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"') {
                    inQuotes = !inQuotes; // 遇到引號，切換狀態 (引號內的逗號不會被切)
                }
                else if (c == ',' && !inQuotes) {
                    result.Add(currentField.ToString());
                    currentField.Clear();
                }
                else {
                    currentField.Append(c);
                }
            }
            result.Add(currentField.ToString());
            return result.ToArray();
        }

        // 🟢 加入 Excel 複製貼上 (Ctrl+V) 防當機事件
        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                try
                {
                    string text = Clipboard.GetText();
                    if (string.IsNullOrEmpty(text)) return;
                    
                    _calcHelper?.BeginBulkUpdate(); // 開始貼上，關閉即時計算

                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex;
                    int c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;

                    foreach (string line in lines)
                    {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++)
                        {
                            if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly)
                            {
                                // 貼上時一併去掉潛在的雙引號
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                            }
                        }
                        r++;
                    }

                    _calcHelper?.EndBulkUpdate(); // 結束貼上，統一計算
                }
                catch (Exception ex)
                {
                    _calcHelper?.EndBulkUpdate();
                    MessageBox.Show(Form.ActiveForm, "貼上失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
