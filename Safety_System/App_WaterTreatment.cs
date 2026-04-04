using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_WaterTreatment
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;

        // 🟢 控制初次載入的開關
        private bool _isFirstLoad = true;

        private const string DbName = "Water"; 
        private const string TableName = "WaterMeterReadings"; 

        public Control GetView()
        {
            // 如果資料庫或資料表不存在，會按原規則在此行自動建立
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WaterMeterReadings] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [廢水處理量] TEXT, 
                [廢水進流量] TEXT, 
                [納廢回收6吋] TEXT, 
                [雙介質A] TEXT, 
                [雙介質B] TEXT, 
                [貯存池] TEXT, 
                [軟水A] TEXT, 
                [軟水B] TEXT, 
                [軟水C] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "廢水處理水量記錄 (庫: Water / 表: WaterMeterReadings)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 25, 10, 10) };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-30) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
            
            Button bRead = new Button { Text = "讀取資料", Size = new Size(120, 35) };
            bRead.Click += (s, e) => RefreshGrid();

            Button bSave = new Button { Text = "💾 儲存", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                if (DataManager.ValidateAndSaveTable(DbName, TableName, dt)) {
                    MessageBox.Show("儲存完成！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshGrid();
                }
            };
            
            Button bExport = new Button { Text = "匯出", Size = new Size(120, 35) };
            bExport.Click += BtnExport_Click;

            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(120, 35) };
            bImport.Click += BtnImportCsv_Click;

            row1.Controls.AddRange(new Control[] { lblRange, _dtpStart, lblTilde, _dtpEnd, bRead, bSave, bExport, bImport });

            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            Label lblOps = new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; 
            _txtNewColName = new TextBox { Width = 120 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) };
            bAdd.Click += (s, e) => {
                if (string.IsNullOrEmpty(_txtNewColName.Text)) return;
                if (VerifyPassword()) {
                    DataManager.AddColumn(DbName, TableName, _txtNewColName.Text);
                    RefreshGrid();
                    _txtNewColName.Clear();
                }
            };

            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "標題更改", Size = new Size(120, 35) };
            bRen.Click += (s, e) => {
                if (_cboColumns.SelectedItem == null || string.IsNullOrEmpty(_txtRenameCol.Text)) return;
                if (VerifyPassword()) {
                    DataManager.RenameColumn(DbName, TableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text);
                    RefreshGrid();
                    _txtRenameCol.Clear();
                }
            };

            Button bDel = new Button { Text = "刪除整列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDel.Click += (s, e) => {
                if (_dgv.CurrentRow == null || _dgv.CurrentRow.Cells["Id"].Value == DBNull.Value) return;
                if (VerifyPassword()) {
                    DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value));
                    RefreshGrid();
                }
            };

            row2.Controls.AddRange(new Control[] { lblOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDel });

            tlp.Controls.Add(row1, 0, 0);
            tlp.Controls.Add(row2, 0, 1);
            boxTop.Controls.Add(tlp);
            main.Controls.Add(boxTop, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);

            // 🟢 在畫面載入完畢後，強制觸發一次讀取
            RefreshGrid();
            return main;
        }

        private void RefreshGrid() {
            if (_isFirstLoad) {
                // 🟢 第一次載入：強迫撈取最新 30 筆
                DataTable dt = DataManager.GetLatestRecords(DbName, TableName, 30);
                _dgv.DataSource = dt;
                _isFirstLoad = false;
                
                // 動態調整日曆區間，自動對齊這 30 筆資料的時間點 (防呆 UX)
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
                // 🟢 之後點擊「讀取資料」，就會依循使用者設定的日期範圍
                _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            }

            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.Add(lbl); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_dgv.Rows.Count == 0 || (_dgv.Rows.Count == 1 && _dgv.Rows[0].IsNewRow)) { 
                MessageBox.Show("沒有資料可匯出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|CSV 檔案 (*.csv)|*.csv", FileName = "廢水處理水量記錄_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) {
                            using (ExcelPackage p = new ExcelPackage()) {
                                var ws = p.Workbook.Worksheets.Add("Data");
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
                                ws.Cells.AutoFitColumns(); 
                                p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        } else {
                            StringBuilder sb = new StringBuilder();
                            string[] colNames = new string[dt.Columns.Count];
                            for (int i = 0; i < dt.Columns.Count; i++) colNames[i] = dt.Columns[i].ColumnName;
                            sb.AppendLine(string.Join(",", colNames));
                            foreach (DataRow row in dt.Rows) {
                                string[] fields = new string[dt.Columns.Count];
                                for (int i = 0; i < dt.Columns.Count; i++) fields[i] = row[i]?.ToString().Replace(",", "，");
                                sb.AppendLine(string.Join(",", fields));
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("資料匯出成功！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案 (*.csv)|*.csv" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length < 2) return; 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) 
                        {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            DataRow nr = dt.NewRow(); 
                            string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) 
                            {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim();
                            }
                            dt.Rows.Add(nr);
                        }
                    } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}
