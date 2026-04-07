using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq; // 支援排序操作
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_DischargeData
    {
        private DataGridView _dgv;
        // 僅保留 年、月 下拉選單 (移除日)
        private ComboBox _cboStartYear, _cboStartMonth;
        private ComboBox _cboEndYear, _cboEndMonth;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        private const string DbName = "Water"; 
        private const string TableName = "DischargeData"; 

        public Control GetView()
        {
            // 初始化資料表 (以「月份」為時間基準)
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [DischargeData] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [月份] TEXT, 
                [水量] TEXT, 
                [SS] TEXT, 
                [COD] TEXT, 
                [BOD] TEXT, 
                [氨氮] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"納管水質排放數據 (庫：{DbName} 表：{TableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            
            // 初始化下拉選單 (年、月)
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 20; i <= currentYear + 10; i++) {
                _cboStartYear.Items.Add(i);
                _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2"));
                _cboEndMonth.Items.Add(i.ToString("D2"));
            }

            // 預設區間：今年 1 月 到 現在
            _cboStartYear.SelectedItem = currentYear;
            _cboStartMonth.SelectedItem = "01";
            _cboEndYear.SelectedItem = currentYear;
            _cboEndMonth.SelectedItem = DateTime.Now.Month.ToString("D2");

            Label lblStartYear = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblStartMonth = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            Label lblEndYear = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEndMonth = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

            Button bRead = new Button { Text = "讀取資料", Size = new Size(110, 35) };
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
                Size = new Size(110, 35), 
                BackColor = Color.ForestGreen, 
                ForeColor = Color.White, 
                Margin = new Padding(10, 3, 3, 3) 
            };
            bSave.Click += (s, e) => {
                _dgv.EndEdit(); 
                SaveColumnOrder(); 
                DataTable dtToSave = (DataTable)_dgv.DataSource;
                EnforceMonthFormat(dtToSave); // 強制 yyyy-MM 格式

                if (DataManager.ValidateAndSaveTable(DbName, TableName, dtToSave)) {
                    MessageBox.Show(Form.ActiveForm, "儲存完成！(已記憶最新欄位排序)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshGrid();
                }
            };
            
            Button bExport = new Button { Text = "匯出", Size = new Size(90, 35) };
            bExport.Click += BtnExport_Click;

            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(110, 35) };
            bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(130, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
                _btnToggle.BackColor = _boxAdvanced.Visible ? Color.LightCoral : Color.LightGray;
            };

            row1.Controls.AddRange(new Control[] { 
                lblRange, _cboStartYear, lblStartYear, _cboStartMonth, lblStartMonth, 
                lblTilde, _cboEndYear, lblEndYear, _cboEndMonth, lblEndMonth, 
                bRead, bExport, bImport, _btnToggle, bSave 
            });
            boxTop.Controls.Add(row1);

            // 進階管理區塊
            _boxAdvanced = new GroupBox { Text = "進階欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdvMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            FlowLayoutPanel row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Label lblOps = new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; 
            _txtNewColName = new TextBox { Width = 120 };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(110, 35) };
            bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && VerifyPassword()) { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };

            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 120 };
            Button bRen = new Button { Text = "標題更改", Size = new Size(110, 35) };
            bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && VerifyPassword()) { DataManager.RenameColumn(DbName, TableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };

            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(110, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null) {
                    string colToDrop = _cboColumns.SelectedItem.ToString();
                    if (MessageBox.Show(Form.ActiveForm, $"警告：確定要刪除整欄【{colToDrop}】嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        if (VerifyPassword()) { DataManager.DropColumn(DbName, TableName, colToDrop); RefreshGrid(); } 
                    }
                } 
            };
            Button bDelRow = new Button { Text = "刪除整列", Size = new Size(110, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += (s, e) => { if (_dgv.CurrentRow != null && _dgv.CurrentRow.Cells["Id"].Value != DBNull.Value && VerifyPassword()) { DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value)); RefreshGrid(); } };
            row2.Controls.AddRange(new Control[] { lblOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });

            FlowLayoutPanel row3 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            Label lblLatest = new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            TextBox txtLatestCount = new TextBox { Width = 100, Text = "30" }; 
            Button bReadLatest = new Button { Text = "讀取筆數", Size = new Size(110, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bReadLatest.Click += (s, e) => {
                if (int.TryParse(txtLatestCount.Text, out int limit) && limit > 0) {
                    DataTable dt = DataManager.GetLatestRecords(DbName, TableName, limit);
                    EnforceMonthFormat(dt);
                    _dgv.DataSource = dt;
                    if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
                    UpdateCboColumns();
                    RestoreColumnOrder();
                    MessageBox.Show(Form.ActiveForm, $"讀取完成！共載入最近 {dt.Rows.Count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            row3.Controls.AddRange(new Control[] { lblLatest, txtLatestCount, bReadLatest });

            flpAdvMain.Controls.Add(row2); flpAdvMain.Controls.Add(row3);
            _boxAdvanced.Controls.Add(flpAdvMain);

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
                EnforceMonthFormat(dt);
                _dgv.DataSource = dt;
                _isFirstLoad = false;
            } else {
                string sDate = $"{_cboStartYear.SelectedItem}-{_cboStartMonth.SelectedItem}";
                string eDate = $"{_cboEndYear.SelectedItem}-{_cboEndMonth.SelectedItem}";
                DataTable dt = DataManager.GetTableData(DbName, TableName, "月份", sDate, eDate);
                EnforceMonthFormat(dt);
                _dgv.DataSource = dt;
            }
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            UpdateCboColumns();
            RestoreColumnOrder();
        }

        private void UpdateCboColumns() {
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "月份") _cboColumns.Items.Add(c.Name);
        }

        private void EnforceMonthFormat(DataTable dt) {
            if (dt == null || !dt.Columns.Contains("月份")) return;
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                string val = row["月份"]?.ToString();
                if (!string.IsNullOrWhiteSpace(val) && DateTime.TryParse(val, out DateTime d)) 
                    row["月份"] = d.ToString("yyyy-MM");
            }
        }

        private void SaveColumnOrder() {
            try {
                var orderedCols = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray();
                File.WriteAllText($"ColOrder_{DbName}_{TableName}.txt", string.Join(",", orderedCols), Encoding.UTF8);
                if (_dgv.DataSource is DataTable dt) {
                    for (int i = 0; i < orderedCols.Length; i++) if (dt.Columns.Contains(orderedCols[i])) dt.Columns[orderedCols[i]].SetOrdinal(i);
                }
            } catch { }
        }

        private void RestoreColumnOrder() {
            try {
                string fileName = $"ColOrder_{DbName}_{TableName}.txt";
                if (File.Exists(fileName)) {
                    string[] savedCols = File.ReadAllText(fileName, Encoding.UTF8).Split(',');
                    for (int i = 0; i < savedCols.Length; i++) if (_dgv.Columns.Contains(savedCols[i])) _dgv.Columns[savedCols[i]].DisplayIndex = i;
                }
            } catch { }
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40 };
            p.Controls.AddRange(new Control[] { new Label { Text = "請輸入管理員密碼：", Left = 30, Top = 30, AutoSize = true }, t, b });
            p.AcceptButton = b;
            return p.ShowDialog(Form.ActiveForm) == DialogResult.OK && t.Text == "tces";
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            if (_dgv.Rows.Count <= 1 && _dgv.Rows[0].IsNewRow) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = "納管排放水質_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog(Form.ActiveForm) == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) { using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); p.SaveAs(new FileInfo(sfd.FileName)); } }
                        else {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows) sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功");
                    } catch (Exception ex) { MessageBox.Show("失敗: " + ex.Message); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV (*.csv)|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) { 
                                string cn = headers[h].Trim(); 
                                if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim(); 
                            }
                            dt.Rows.Add(nr);
                        }
                        MessageBox.Show($"已載入 {lines.Length - 1} 筆，請按儲存。");
                    } catch (Exception ex) { MessageBox.Show("匯入失敗: " + ex.Message); }
                }
            }
        }
    }
}
