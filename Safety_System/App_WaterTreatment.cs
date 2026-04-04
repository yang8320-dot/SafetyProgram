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

        private const string DbName = "Water"; 
        private const string TableName = "WaterMeterReadings"; 

        public Control GetView()
        {
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

            // 🟢 修正：主排版 Padding 加入 Top 40，將整個控制區往下移
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                RowCount = 2,
                Padding = new Padding(15, 40, 15, 10) 
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 🟢 修正：boxTop 內邊距，增加內部空間感
            GroupBox boxTop = new GroupBox { 
                Text = "水處理數據管理 (庫: " + DbName + " / 表: " + TableName + ")", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                AutoSize = true, 
                Padding = new Padding(15, 30, 15, 15)
            };

            TableLayoutPanel tlpControls = new TableLayoutPanel { 
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true 
            };

            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            flpRow1.Controls.Add(new Label { Text = "日期起:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _dtpStart = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Margin = new Padding(3, 5, 3, 3) };
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            flpRow1.Controls.Add(_dtpStart);
            
            flpRow1.Controls.Add(new Label { Text = "日期迄:", AutoSize = true, Margin = new Padding(15, 10, 3, 0) });
            _dtpEnd = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Margin = new Padding(3, 5, 3, 3) };
            flpRow1.Controls.Add(_dtpEnd);
            
            Button btnRead = new Button { Text = "讀取資料庫", Size = new Size(110, 35), BackColor = Color.LightBlue, Margin = new Padding(15, 2, 3, 3) };
            btnRead.Click += (s, e) => RefreshGrid();
            flpRow1.Controls.Add(btnRead);

            Button btnImportCsv = new Button { Text = "📥 匯入 CSV", Size = new Size(110, 35), BackColor = Color.Orange, Margin = new Padding(5, 2, 3, 3) };
            btnImportCsv.Click += BtnImportCsv_Click;
            flpRow1.Controls.Add(btnImportCsv);

            Button btnSave = new Button { 
                Text = "💾 儲存", Size = new Size(100, 35), 
                BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(5, 2, 3, 3)
            };
            btnSave.Click += BtnSave_Click;
            flpRow1.Controls.Add(btnSave);

            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 15, 0, 5) };
            flpRow2.Controls.Add(new Label { Text = "新增欄位:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _txtNewColName = new TextBox { Width = 150, Margin = new Padding(3, 7, 3, 3) }; 
            flpRow2.Controls.Add(_txtNewColName);
            Button btnAddCol = new Button { Text = "確認新增", Size = new Size(100, 35), BackColor = Color.LightGray, Margin = new Padding(5, 2, 3, 3) };
            btnAddCol.Click += (s, e) => {
                if (!string.IsNullOrWhiteSpace(_txtNewColName.Text)) {
                    DataManager.AddColumn(DbName, TableName, _txtNewColName.Text.Trim());
                    _txtNewColName.Clear(); RefreshGrid();
                }
            };
            flpRow2.Controls.Add(btnAddCol);

            flpRow2.Controls.Add(new Label { Text = "管理欄位:", AutoSize = true, Margin = new Padding(25, 10, 3, 0) });
            _cboColumns = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3, 7, 3, 3) };
            flpRow2.Controls.Add(_cboColumns);
            _txtRenameCol = new TextBox { Width = 140, Margin = new Padding(5, 7, 3, 3) }; 
            flpRow2.Controls.Add(_txtRenameCol);
            
            Button btnRenameCol = new Button { Text = "✏️ 改名", Size = new Size(80, 35), BackColor = Color.LightYellow, Margin = new Padding(5, 2, 3, 3) };
            btnRenameCol.Click += BtnRenameCol_Click;
            flpRow2.Controls.Add(btnRenameCol);
            
            Button btnDropCol = new Button { Text = "⚠️ 刪除欄", Size = new Size(100, 35), BackColor = Color.LightPink, Margin = new Padding(5, 2, 3, 3) };
            btnDropCol.Click += BtnDropCol_Click;
            flpRow2.Controls.Add(btnDropCol);

            Button btnDeleteRow = new Button {
                Text = "🗑️ 刪除選取資料", Size = new Size(150, 35), 
                BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(15, 2, 3, 3)
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;
            flpRow2.Controls.Add(btnDeleteRow);

            tlpControls.Controls.Add(flpRow1, 0, 0);
            tlpControls.Controls.Add(flpRow2, 0, 1);
            boxTop.Controls.Add(tlpControls);
            mainLayout.Controls.Add(boxTop, 0, 0);

            GroupBox boxBottom = new GroupBox { Text = "數據明細 (支援 Ctrl+V 貼上、右鍵匯出)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToAddRows = true
            };
            _dgv.KeyDown += Dgv_KeyDown;
            
            // 🟢 修正：自動預填今日日期
            _dgv.DefaultValuesNeeded += (s, e) => {
                e.Row.Cells["日期"].Value = DateTime.Now.ToString("yyyy-MM-dd");
            };

            SetupContextMenu();
            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        private void RefreshGrid()
        {
            // 🟢 修正：讀取時確保時間字串格式化，避免 SQLite 字串比對失敗
            string start = _dtpStart.Value.ToString("yyyy-MM-dd");
            string end = _dtpEnd.Value.ToString("yyyy-MM-dd");
            
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", start, end);
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            
            if (_dgv.Columns.Contains("日期"))
            {
                _dgv.Columns["日期"].DefaultCellStyle.Format = "yyyy-MM-dd";
            }

            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col.Name != "Id" && col.Name != "日期") _cboColumns.Items.Add(col.Name);
            }
            if (_cboColumns.Items.Count > 0) _cboColumns.SelectedIndex = 0;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // 🟢 修正：存檔前強制結束 DataGridView 編輯，確保最後一筆資料有提交
            _dgv.EndEdit();
            if (_dgv.BindingContext != null && _dgv.DataSource != null)
            {
                _dgv.BindingContext[_dgv.DataSource].EndCurrentEdit();
            }

            DataTable dt = (DataTable)_dgv.DataSource;
            if (dt == null) return;
            try {
                foreach (DataRow row in dt.Rows) {
                    if (row.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, row);
                }
                MessageBox.Show("資料儲存完畢！"); RefreshGrid();
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
        }

        private void BtnDeleteRow_Click(object sender, EventArgs e)
        {
            if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
            var idVal = _dgv.CurrentRow.Cells["Id"].Value;
            if (idVal != null && idVal != DBNull.Value) {
                if (MessageBox.Show("確定刪除此筆資料？", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(idVal));
                    RefreshGrid();
                }
            } else { _dgv.Rows.Remove(_dgv.CurrentRow); }
        }

        private void BtnRenameCol_Click(object sender, EventArgs e)
        {
            string targetCol = _cboColumns.Text;
            string newName = _txtRenameCol.Text.Trim();
            if (string.IsNullOrEmpty(targetCol) || string.IsNullOrEmpty(newName)) return;
            if (!VerifyPassword()) return;
            try {
                DataManager.RenameColumn(DbName, TableName, targetCol, newName);
                MessageBox.Show("更名成功！"); _txtRenameCol.Clear(); RefreshGrid();
            } catch (Exception ex) { MessageBox.Show("失敗：" + ex.Message); }
        }

        private void BtnDropCol_Click(object sender, EventArgs e)
        {
            string targetCol = _cboColumns.Text;
            if (string.IsNullOrEmpty(targetCol)) return;
            if (MessageBox.Show($"確定徹底刪除 [{targetCol}] 欄位？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                if (!VerifyPassword()) return;
                try { DataManager.DropColumn(DbName, TableName, targetCol); RefreshGrid(); }
                catch (Exception ex) { MessageBox.Show("失敗：" + ex.Message); }
            }
        }

        private bool VerifyPassword()
        {
            Form prompt = new Form() { Width = 450, Height = 240, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "安全授權驗證", StartPosition = FormStartPosition.CenterParent };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "執行此操作需要管理員權限，\n請輸入授權密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 14F) };
            TextBox txt = new TextBox() { Left = 30, Top = 95, Width = 370, PasswordChar = '*', Font = new Font("Microsoft JhengHei UI", 14F) };
            Button btn = new Button() { Text = "確認執行", Left = 280, Top = 145, Width = 120, Height = 40, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            prompt.Controls.Add(lbl); prompt.Controls.Add(txt); prompt.Controls.Add(btn);
            prompt.AcceptButton = btn;
            return prompt.ShowDialog() == DialogResult.OK && txt.Text == "tces";
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) { if (e.Control && e.KeyCode == Keys.V) PasteClipboard(); }

        private void PasteClipboard()
        {
            try {
                string text = Clipboard.GetText();
                if (string.IsNullOrEmpty(text)) return;
                if (_dgv.IsCurrentCellInEditMode) _dgv.EndEdit();
                
                string[] lines = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (string line in lines) {
                    if (string.IsNullOrEmpty(line)) continue;
                    if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                    string[] cells = line.Split('\t');
                    for (int i = 0; i < cells.Length; i++) {
                        if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly)
                        {
                            string val = cells[i].Trim();
                            // 🟢 修正：貼上時若是日期欄位，自動格式化
                            if (_dgv.Columns[c + i].Name == "日期" && DateTime.TryParse(val, out DateTime dateVal))
                            {
                                val = dateVal.ToString("yyyy-MM-dd");
                            }
                            _dgv[c + i, r].Value = val;
                        }
                    }
                    r++;
                }
            } catch (Exception ex) { MessageBox.Show("貼上失敗：" + ex.Message); }
        }

        private void SetupContextMenu() {
            ContextMenuStrip cms = new ContextMenuStrip();
            cms.Items.Add("📤 匯出 CSV", null, (s, e) => ExportData("csv"));
            cms.Items.Add("📊 匯出 XLSX", null, (s, e) => ExportData("xlsx"));
            _dgv.ContextMenuStrip = cms;
        }

        private void ExportData(string format)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = format == "csv" ? "CSV|*.csv" : "Excel|*.xlsx" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (format == "csv") {
                            StringBuilder sb = new StringBuilder();
                            foreach (DataColumn col in dt.Columns) sb.Append(col.ColumnName + ",");
                            sb.AppendLine();
                            foreach (DataRow row in dt.Rows) {
                                foreach (var item in row.ItemArray) sb.Append(item.ToString().Replace(",", "，") + ",");
                                sb.AppendLine();
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                        } else {
                            using (var p = new ExcelPackage()) {
                                var ws = p.Workbook.Worksheets.Add("Data");
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
                                ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        }
                        MessageBox.Show("匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) {
                            DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") {
                                    string val = vs[h].Trim();
                                    if (cn == "日期" && DateTime.TryParse(val, out DateTime d)) val = d.ToString("yyyy-MM-dd");
                                    nr[cn] = val;
                                }
                            }
                            dt.Rows.Add(nr);
                        }
                        RefreshGrid();
                    } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message); }
                }
            }
        }
    }
}
