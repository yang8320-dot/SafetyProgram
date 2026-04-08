/// FILE: Safety_System/App_Law_Generic.cs ///
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_Law_Generic
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        
        // 接收外部傳入的資料庫與資料表名稱
        private readonly string _dbName; 
        private readonly string _tableName; 

        public App_Law_Generic(string dbName, string tableName)
        {
            _dbName = dbName;
            _tableName = tableName;
        }

        public Control GetView()
        {
            // 動態初始化資料表
            DataManager.InitTable(_dbName, _tableName, $@"CREATE TABLE IF NOT EXISTS [{_tableName}] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [法規名稱] TEXT, 
                [發布機關] TEXT, 
                [施行日期] TEXT, 
                [條] TEXT,
                [項] TEXT,
                [款] TEXT,
                [目] TEXT,
                [內容] TEXT,
                [重點摘要] TEXT, 
                [適用性] TEXT, 
                [鑑別日期] TEXT,
                [備註] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"法規管理 (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Label lblRange = new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddYears(-1) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
            
            Button bRead = new Button { Text = "讀取資料", Size = new Size(150, 35) };
            bRead.Click += (s, e) => { RefreshGrid(); if (!_isFirstLoad) { int count = ((DataTable)_dgv.DataSource).Rows.Count; MessageBox.Show(Form.ActiveForm, $"讀取完成！共找到 {count} 筆資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); } };

            Button bSave = new Button { Name = "btnSave", Text = "💾 儲存", Size = new Size(150, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Margin = new Padding(30, 0, 0, 0) };
            bSave.Click += (s, e) => { _dgv.EndEdit(); if (DataManager.ValidateAndSaveTable(_dbName, _tableName, (DataTable)_dgv.DataSource)) { MessageBox.Show(Form.ActiveForm, "儲存完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); RefreshGrid(); } };
            
            Button bExport = new Button { Text = "匯出 Excel", Size = new Size(150, 35) }; bExport.Click += BtnExport_Click;
            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(150, 35) }; bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => { _boxAdvanced.Visible = !_boxAdvanced.Visible; _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理"; _btnToggle.BackColor = _boxAdvanced.Visible ? Color.LightCoral : Color.LightGray; };

            row1.Controls.AddRange(new Control[] { lblRange, _dtpStart, lblTilde, _dtpEnd, bRead, bExport, bImport, _btnToggle, bSave });
            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdvMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            FlowLayoutPanel row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Label lblOps = new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }; _txtNewColName = new TextBox { Width = 120 };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) }; bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && VerifyPassword()) { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            Button bRen = new Button { Text = "標題更改", Size = new Size(120, 35) }; bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && VerifyPassword()) { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { if (_cboColumns.SelectedItem != null) { string colToDrop = _cboColumns.SelectedItem.ToString(); if (MessageBox.Show(Form.ActiveForm, $"警告：確定要刪除整欄【{colToDrop}】嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) { if (VerifyPassword()) { DataManager.DropColumn(_dbName, _tableName, colToDrop); RefreshGrid(); } } } };
            Button bDelRow = new Button { Text = "刪除整列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White }; bDelRow.Click += (s, e) => { if (_dgv.CurrentRow != null && _dgv.CurrentRow.Cells["Id"].Value != DBNull.Value && VerifyPassword()) { DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value)); RefreshGrid(); } };
            row2.Controls.AddRange(new Control[] { lblOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });

            flpAdvMain.Controls.Add(row2); _boxAdvanced.Controls.Add(flpAdvMain);

            // 🟢 DataGridView 初始化：開啟列高自動延展功能
            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells // 讓高度適應文字行數
            };
            
            // 🟢 攔截 DataError 與 編輯控制項事件 (這兩個是防呆與換行的關鍵)
            _dgv.DataError += Dgv_DataError;
            _dgv.EditingControlShowing += Dgv_EditingControlShowing;

            main.Controls.Add(boxTop, 0, 0); main.Controls.Add(_boxAdvanced, 0, 1); main.Controls.Add(_dgv, 0, 2);
            RefreshGrid(); return main;
        }

        private void RefreshGrid() {
            if (_isFirstLoad) {
                DataTable dt = DataManager.GetLatestRecords(_dbName, _tableName, 50); 
                _dgv.DataSource = dt;
                _isFirstLoad = false;
            } else { 
                _dgv.DataSource = DataManager.GetTableData(_dbName, _tableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd")); 
            }
            
            // 🟢 呼叫轉換下拉選單的方法
            SetupComboBoxColumns();
            
            // 🟢 呼叫設定長文字自動換行與寬度限制的方法
            SetupTextWrapping();

            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true; 
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
            
            // 強制重繪表格列高
            _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        }

        // 🟢 將指定欄位替換為下拉選單
        private void SetupComboBoxColumns()
        {
            ReplaceWithComboBox("適用性", new string[] { "適用", "不適用", "參考", "確認中", "" });

            string[] itemsTiao = new string[501];
            itemsTiao[0] = "";
            for (int i = 1; i <= 500; i++) itemsTiao[i] = i.ToString();
            ReplaceWithComboBox("條", itemsTiao);

            string[] itemsSmall = new string[21];
            itemsSmall[0] = "";
            for (int i = 1; i <= 20; i++) itemsSmall[i] = i.ToString();
            
            ReplaceWithComboBox("項", itemsSmall);
            ReplaceWithComboBox("款", itemsSmall);
            ReplaceWithComboBox("目", itemsSmall);
        }

        private void ReplaceWithComboBox(string colName, string[] items)
        {
            if (_dgv.Columns.Contains(colName) && !(_dgv.Columns[colName] is DataGridViewComboBoxColumn))
            {
                int colIndex = _dgv.Columns[colName].Index;
                _dgv.Columns.Remove(colName);

                DataGridViewComboBoxColumn cboCol = new DataGridViewComboBoxColumn();
                cboCol.Name = colName;
                cboCol.HeaderText = colName;
                cboCol.DataPropertyName = colName; 
                cboCol.Items.AddRange(items);
                
                cboCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                cboCol.FlatStyle = FlatStyle.Flat;
                cboCol.SortMode = DataGridViewColumnSortMode.Automatic; 

                _dgv.Columns.Insert(colIndex, cboCol);
            }
        }

        // 🟢 鎖定長文字欄位的寬度並開啟自動換行
        private void SetupTextWrapping()
        {
            string[] longTextColumns = { "法規名稱", "內容", "重點摘要", "備註" };

            foreach (string colName in longTextColumns)
            {
                if (_dgv.Columns.Contains(colName))
                {
                    // 1. 強制關閉單一欄位的自動撐寬，否則 WinForm 預設會一直延伸為一行
                    _dgv.Columns[colName].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    
                    // 2. 給予固定的顯示寬度 (若覺得不夠寬可自行把 350 改大)
                    _dgv.Columns[colName].Width = 350; 
                    
                    // 3. 啟用文字換行屬性
                    _dgv.Columns[colName].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }
        }

        // 🟢 控制編輯狀態 (包含: 下拉選單防呆、文字框支援換行)
        private void Dgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cbo)
            {
                // 【下拉選單防呆】設定為 DropDownList 模式，確保使用者只能從清單選取，不能打字
                cbo.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            else if (e.Control is TextBox txt)
            {
                // 【文字框防呆】解決填表時文字會無限向右延伸的問題，開啟多行模式讓游標自動折行
                txt.Multiline = true;
            }
        }

        // 🟢 忽略選單資料比對錯誤 (舊資料不在選單內時不當機)
        private void Dgv_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.AddRange(new Control[] { new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) }, t, b }); p.AcceptButton = b;
            return p.ShowDialog(Form.ActiveForm) == DialogResult.OK && t.Text == "tces";
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            if (_dgv.Rows.Count <= 1) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = _tableName + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable((DataTable)_dgv.DataSource, true); ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName)); }
                    MessageBox.Show("資料匯出成功！");
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV (*.csv)|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default); if (lines.Length < 2) return; 
                    DataTable dt = (DataTable)_dgv.DataSource; string[] headers = lines[0].Split(',');
                    for (int i = 1; i < lines.Length; i++) {
                        if (string.IsNullOrWhiteSpace(lines[i])) continue; DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                        for (int h = 0; h < headers.Length && h < vs.Length; h++) { string cn = headers[h].Trim(); if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim(); }
                        dt.Rows.Add(nr);
                    }
                    MessageBox.Show($"載入 {lines.Length - 1} 筆！請記得按儲存。");
                }
            }
        }
    }
}
