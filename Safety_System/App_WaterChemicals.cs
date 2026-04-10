/// FILE: Safety_System/App_WaterChemicals.cs ///
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
    public class App_WaterChemicals
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
        private const string TableName = "WaterChemicals"; 
        
        private DataGridViewAutoCalcHelper _calcHelper; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WaterChemicals] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [星期] TEXT,
                [PAC] TEXT, 
                [PAC日統計] TEXT, 
                [NAOH] TEXT, 
                [NAOH日統計] TEXT, 
                [高分子] TEXT, 
                [高分子日統計] TEXT, 
                [備註] TEXT);");

            if (!DataManager.GetColumnNames(DbName, TableName).Contains("星期"))
            {
                DataManager.AddColumn(DbName, TableName, "星期");
            }

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"廢水處理用藥記錄 (庫：{DbName} 表：{TableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            
            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

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

            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            Button bRead = new Button { Text = "🔍 讀取資料", Size = new Size(150, 35), BackColor = Color.WhiteSmoke };
            bRead.Click += (s, e) => { RefreshGrid(); if (!_isFirstLoad) MessageBox.Show("資料載入完成！"); };

            Button bSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(150, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            bSave.Click += BtnSave_Click; 
            
            Button bExport = new Button { Text = "📤 匯出Excel", Size = new Size(150, 35) }; bExport.Click += BtnExport_Click;
            Button bImport = new Button { Text = "📥 匯入CSV", Size = new Size(150, 35) }; bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
            };

            row1.Controls.Add(lblRange);
            row1.Controls.Add(_cboStartYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboStartMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboStartDay); row1.Controls.Add(new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) });
            row1.Controls.Add(_cboEndYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboEndMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboEndDay); row1.Controls.Add(new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(bRead); row1.Controls.Add(bExport); row1.Controls.Add(bImport); row1.Controls.Add(_btnToggle); row1.Controls.Add(bSave);
            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true };
            _txtNewColName = new TextBox { Width = 150 };
            
            // 🟢 管理員權限
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(100, 35) };
            bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            
            // 🟢 管理員權限
            Button bRen = new Button { Text = "修改名稱", Size = new Size(100, 35) };
            bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { DataManager.RenameColumn(DbName, TableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };
            
            // 🟢 管理員權限
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(100, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { DataManager.DropColumn(DbName, TableName, _cboColumns.SelectedItem.ToString()); RefreshGrid(); } };
            
            // 🟢 一般權限 + 複選刪除邏輯
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                                       .Select(c => c.OwningRow)
                                       .Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value)
                                       .Distinct().ToList();

                if (selectedRows.Count > 0) {
                    if (MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(刪除後將立即生效)", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        if (AuthManager.VerifyUser()) {
                            foreach (var r in selectedRows) {
                                DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(r.Cells["Id"].Value));
                            }
                            RefreshGrid();
                            MessageBox.Show("刪除成功！");
                        }
                    }
                } else {
                    MessageBox.Show("請先用滑鼠選取要刪除的資料列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += (s, e) => { if (int.TryParse(txtLimit.Text, out int l)) { DataTable dt = DataManager.GetLatestRecords(DbName, TableName, l); EnforceMonthFormat(dt); _dgv.DataSource = dt; RestoreColumnOrder(); } };
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
            flpAdv.Controls.Add(rowAdv1); flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToOrderColumns = true };
            _dgv.KeyDown += Dgv_KeyDown;
            
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); main.Controls.Add(_boxAdvanced, 0, 1); main.Controls.Add(_dgv, 0, 2);

            RefreshGrid();
            return main;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit();
                SaveColumnOrder();
                DataTable dt = (DataTable)_dgv.DataSource;
                EnforceMonthFormat(dt);
                if (DataManager.BulkSaveTable(DbName, TableName, dt)) {
                    MessageBox.Show("儲存完成！");
                    RefreshGrid();
                }
            } finally {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void RefreshGrid()
        {
            DataTable dt;
            if (_isFirstLoad) {
                dt = DataManager.GetLatestRecords(DbName, TableName, 30);
                _isFirstLoad = false;
            } else {
                string sDate = GetStartDate().ToString("yyyy-MM-dd");
                string eDate = GetEndDate().ToString("yyyy-MM-dd");
                dt = DataManager.GetTableData(DbName, TableName, "日期", sDate, eDate);
            }
            EnforceMonthFormat(dt);
            _dgv.DataSource = dt;
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            UpdateCboColumns();
            RestoreColumnOrder();
        }

        private void UpdateCboColumns()
        {
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns)
                if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private void EnforceMonthFormat(DataTable dt)
        {
            if (dt == null || !dt.Columns.Contains("月份")) return;
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                if (DateTime.TryParse(row["月份"]?.ToString(), out DateTime d)) row["月份"] = d.ToString("yyyy-MM");
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetStartDate() => ParseComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
        private DateTime GetEndDate() => ParseComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);
        private DateTime ParseComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime def) { try { return new DateTime(int.Parse(y.SelectedItem.ToString()), int.Parse(m.SelectedItem.ToString()), int.Parse(d.SelectedItem.ToString())); } catch { return def; } }

        private void SaveColumnOrder() { try { var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); File.WriteAllText($"ColOrder_{DbName}_{TableName}.txt", string.Join(",", ordered), Encoding.UTF8); } catch { } }
        private void RestoreColumnOrder() { try { string fn = $"ColOrder_{DbName}_{TableName}.txt"; if (File.Exists(fn)) { string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); for (int i = 0; i < saved.Length; i++) if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; } } catch { } }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = TableName + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) {
                            using (ExcelPackage p = new ExcelPackage()) {
                                var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        } else {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows) sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案 (*.csv)|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length < 2) return;
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = ParseCsvLine(lines[0]);

                        _dgv.DataSource = null; 
                        _calcHelper?.BeginBulkUpdate();

                        foreach (string line in lines.Skip(1)) {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            DataRow nr = dt.NewRow();
                            string[] vs = ParseCsvLine(line);
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim().Trim('"');
                            }
                            dt.Rows.Add(nr);
                        }

                        _calcHelper?.RecalculateTable(dt); 
                        _calcHelper?.EndBulkUpdate();
                        _dgv.DataSource = dt; 
                        RestoreColumnOrder();
                        MessageBox.Show("匯入成功!請檢查數據後點擊儲存。");
                    } catch (Exception ex) { RefreshGrid(); MessageBox.Show("匯入異常：" + ex.Message); }
                }
            }
        }

        private string[] ParseCsvLine(string line) { var res = new System.Collections.Generic.List<string>(); bool q = false; var f = new StringBuilder(); foreach (char c in line) { if (c == '\"') q = !q; else if (c == ',' && !q) { res.Add(f.ToString()); f.Clear(); } else f.Append(c); } res.Add(f.ToString()); return res.ToArray(); }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V) {
                try {
                    string text = Clipboard.GetText(); if (string.IsNullOrEmpty(text)) return;
                    _calcHelper?.BeginBulkUpdate();
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;
                    foreach (string line in lines) {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++)
                            if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly)
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt);
                    _calcHelper?.EndBulkUpdate();
                } catch { _calcHelper?.EndBulkUpdate(); }
            }
        }
    }
}
