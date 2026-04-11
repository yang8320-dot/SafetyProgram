/// FILE: Safety_System/App_Test_Generic.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq; // 🟢 補上 LINQ 引用，解決 Cast 和 Select 報錯
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_Test_Generic
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        
        // 動態傳入的參數
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;

        // 集中管理各個檢測表的初始欄位架構 (Schema)
        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "EnvMonitor", "[日期] TEXT, [測點名稱] TEXT, [溫度] TEXT, [濕度] TEXT, [噪音(dB)] TEXT, [照度(Lux)] TEXT, [備註] TEXT" },
            { "WastewaterPeriodic", "[日期] TEXT, [申報季別] TEXT, [排放水量] TEXT, [COD] TEXT, [SS] TEXT, [BOD] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "DrinkingWater", "[日期] TEXT, [採樣點位置] TEXT, [大腸桿菌群] TEXT, [總菌落數] TEXT, [鉛] TEXT, [濁度] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "IndustrialZoneTest", "[日期] TEXT, [採樣點位置] TEXT, [水溫] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [重金屬] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "SoilGasTest", "[日期] TEXT, [採樣井編號] TEXT, [測漏氣體濃度] TEXT, [甲烷] TEXT, [二氧化碳] TEXT, [氧氣] TEXT, [檢測機構] TEXT, [備註] TEXT" },
            { "WastewaterSelfTest", "[日期] TEXT, [採樣時間] TEXT, [採樣位置] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [透視度] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "CoolingWaterVendor", "[日期] TEXT, [廠商名稱] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [添加藥劑] TEXT, [檢驗結果] TEXT, [備註] TEXT" },
            { "CoolingWaterSelf", "[日期] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "TCLP", "[日期] TEXT, [樣品名稱] TEXT, [鎘] TEXT, [鉛] TEXT, [鉻] TEXT, [砷] TEXT, [銅] TEXT, [鋅] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "WaterMeterCalibration", "[日期] TEXT, [水錶編號] TEXT, [水錶位置] TEXT, [校正前讀數] TEXT, [校正後讀數] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [備註] TEXT" },
            { "OtherTests", "[日期] TEXT, [檢測項目] TEXT, [檢測位置] TEXT, [檢測數值] TEXT, [單位] TEXT, [合格標準] TEXT, [檢測機構] TEXT, [備註] TEXT" }
        };

        public App_Test_Generic(string dbName, string tableName, string chineseTitle)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
        }

        public Control GetView()
        {
            string schema = _schemaMap.ContainsKey(_tableName) ? _schemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-30) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
            
            Button bRead = new Button { Text = "讀取資料", Size = new Size(120, 35) };
            bRead.Click += (s, e) => { RefreshGrid(); if (!_isFirstLoad) { MessageBox.Show(Form.ActiveForm, $"讀取完成！共找到 {((DataTable)_dgv.DataSource).Rows.Count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); } };

            Button bSave = new Button { Name = "btnSave", Text = "💾 儲存", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Margin = new Padding(30, 0, 0, 0) };
            bSave.Click += (s, e) => { _dgv.EndEdit(); if (DataManager.ValidateAndSaveTable(_dbName, _tableName, (DataTable)_dgv.DataSource)) { MessageBox.Show(Form.ActiveForm, "儲存完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); RefreshGrid(); } };
            
            Button bExport = new Button { Text = "匯出", Size = new Size(120, 35) }; bExport.Click += BtnExport_Click;
            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(120, 35) }; bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => { _boxAdvanced.Visible = !_boxAdvanced.Visible; _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理"; _btnToggle.BackColor = _boxAdvanced.Visible ? Color.LightCoral : Color.LightGray; };

            row1.Controls.AddRange(new Control[] { lblRange, _dtpStart, lblTilde, _dtpEnd, bRead, bExport, bImport, _btnToggle, bSave });
            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdvMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            FlowLayoutPanel row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Label lblOps = new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; _txtNewColName = new TextBox { Width = 120 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) }; 
            bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };
            
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "標題更改", Size = new Size(120, 35) }; 
            bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { if (MessageBox.Show($"確定要刪除整欄【{_cboColumns.SelectedItem}】嗎？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) { DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); RefreshGrid(); } } };
            
            Button bDelRow = new Button { Text = "刪除整列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White }; 
            bDelRow.Click += (s, e) => { if (_dgv.CurrentRow != null && _dgv.CurrentRow.Cells["Id"].Value != DBNull.Value && AuthManager.VerifyUser()) { DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value)); RefreshGrid(); } };
            
            row2.Controls.AddRange(new Control[] { lblOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });

            FlowLayoutPanel row3 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            Label lblLatest = new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; TextBox txtLatestCount = new TextBox { Width = 120, Text = "50" }; 
            Button bReadLatest = new Button { Text = "讀取筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bReadLatest.Click += (s, e) => {
                if (int.TryParse(txtLatestCount.Text, out int limit) && limit > 0) {
                    DataTable dt = DataManager.GetLatestRecords(_dbName, _tableName, limit); _dgv.DataSource = dt;
                    if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true; 
                    _cboColumns.Items.Clear(); foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
                    if (dt.Rows.Count > 0) {
                        DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                        foreach(DataRow r in dt.Rows) { if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { if (d < minD) minD = d; if (d > maxD) maxD = d; } }
                        if (minD <= maxD) { _dtpStart.Value = minD; _dtpEnd.Value = maxD; }
                    }
                }
            };
            row3.Controls.AddRange(new Control[] { lblLatest, txtLatestCount, bReadLatest });
            flpAdvMain.Controls.Add(row2); flpAdvMain.Controls.Add(row3); _boxAdvanced.Controls.Add(flpAdvMain);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            _dgv.KeyDown += Dgv_KeyDown;

            main.Controls.Add(boxTop, 0, 0); main.Controls.Add(_boxAdvanced, 0, 1); main.Controls.Add(_dgv, 0, 2);
            RefreshGrid(); return main;
        }

        private void RefreshGrid() {
            if (_isFirstLoad) {
                DataTable dt = DataManager.GetLatestRecords(_dbName, _tableName, 30); _dgv.DataSource = dt;
                if (dt.Rows.Count > 0) {
                    DateTime minD = DateTime.MaxValue, maxD = DateTime.MinValue;
                    foreach(DataRow r in dt.Rows) if (DateTime.TryParse(r["日期"]?.ToString(), out DateTime d)) { if (d < minD) minD = d; if (d > maxD) maxD = d; }
                    if (minD <= maxD) { _dtpStart.Value = minD; _dtpEnd.Value = maxD; }
                }
                _isFirstLoad = false;
            } else { 
                _dgv.DataSource = DataManager.GetTableData(_dbName, _tableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd")); 
            }
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true; 
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            if (_dgv.Rows.Count == 0 || (_dgv.Rows.Count == 1 && _dgv.Rows[0].IsNewRow)) { MessageBox.Show("沒有資料可匯出！"); return; }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = _chineseTitle + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) { using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName)); } }
                        else {
                            StringBuilder sb = new StringBuilder(); sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows) sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，")))); File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV (*.csv)|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default); if (lines.Length < 2) return; 
                        DataTable dt = (DataTable)_dgv.DataSource; string[] headers = lines[0].Split(',');
                        _dgv.DataSource = null; 
                        for (int i = 1; i < lines.Length; i++) {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue; DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) { string cn = headers[h].Trim(); if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim().Trim('"'); }
                            dt.Rows.Add(nr);
                        }
                        _dgv.DataSource = dt;
                        MessageBox.Show($"載入 {lines.Length - 1} 筆成功！請記得按儲存。");
                    } catch (Exception ex) { RefreshGrid(); MessageBox.Show("匯入失敗：" + ex.Message); }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.V) {
                try {
                    string text = Clipboard.GetText(); if (string.IsNullOrEmpty(text)) return;
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex; DataTable dt = (DataTable)_dgv.DataSource;
                    foreach (string line in lines) {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow()); string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++) if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly) _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                        r++;
                    }
                } catch { }
            }
        }
    }
}
