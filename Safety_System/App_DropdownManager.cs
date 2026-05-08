/// FILE: Safety_System/App_DropdownManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_DropdownManager : Form
    {
        private ComboBox _cboDb, _cboTable;
        private TextBox[] _txtOptions;
        private ComboBox[] _cboCols;
        private ComboBox[] _cboParentVals;
        private Button _btnSave, _btnExport, _btnImport;
        
        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        private readonly Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        public App_DropdownManager()
        {
            try {
                _dbMap = App_DbConfig.GetDbMapCache();
                InitializeComponent();
                LoadDropdownConfigs();
            } catch (Exception ex) {
                MessageBox.Show($"初始化連動選單管理介面時發生嚴重錯誤：\n{ex.Message}\n{ex.StackTrace}", "系統崩潰防護", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "下拉選單與連動項目管理";
            this.Size = new Size(1650, 900);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            // ================= 頂部操作區 =================
            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            Label lblTitle = new Label { Text = "🔧 下拉選單與連動項目設定", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            
            Label lblDb = new Label { Text = "選擇資料庫：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboDb = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 30, 0) };
            
            Label lblTable = new Label { Text = "選擇資料表：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboTable = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 40, 0) };

            _btnExport = new Button { Text = "📤 匯出 Excel", Size = new Size(150, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 15, 0) };
            _btnImport = new Button { Text = "📥 匯入 Excel", Size = new Size(150, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 0) };

            flpControls.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, _btnExport, _btnImport });
            flpTopMain.Controls.Add(lblTitle);
            flpTopMain.Controls.Add(flpControls);
            pnlTop.Controls.Add(flpTopMain);
            this.Controls.Add(pnlTop);

            // ================= 底部儲存區 =================
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 90, BackColor = Color.White, Padding = new Padding(20, 15, 20, 15) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSave = new Button { Text = "💾 儲存並套用當前設定", Dock = DockStyle.Right, Width = 250, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSave.Click += BtnSave_Click;

            Label lblHint = new Label { Text = "※ 第一層的選項修改後，請先點擊【儲存】，再於第二層的「觸發條件」中選取，以設定對應的連動清單。\n※ 選項內容的排列順序，即為系統表單中下拉選單顯示的順序。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0, 5, 0, 0) };

            pnlBottom.Controls.Add(lblHint);
            pnlBottom.Controls.Add(_btnSave);
            this.Controls.Add(pnlBottom);

            // ================= 四層連動編輯區 (完美防護越界與破版) =================
            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            for(int i = 0; i < 4; i++) tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            _cboCols = new ComboBox[4];
            _cboParentVals = new ComboBox[4];
            _txtOptions = new TextBox[4];

            string[] headers = { "第一層 (主選項)", "第二層 (依第一層連動)", "第三層 (依第二層連動)", "第四層 (依第三層連動)" };

            for (int i = 0; i < 4; i++)
            {
                // 底板：極緊湊設計
                Panel pCol = new Panel { 
                    Dock = DockStyle.Fill, 
                    Margin = new Padding(3, 0, 3, 0), 
                    BackColor = Color.White, 
                    Padding = new Padding(15) 
                };
                pCol.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pCol.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                // 固定控制區
                Panel pTopControls = new Panel { Dock = DockStyle.Top, Height = 195, BackColor = Color.White };

                Label lHeader = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Location = new Point(0, 0) };
                
                Label lCol = new Label { Text = "綁定資料表欄位：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 40) };
                _cboCols[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(0, 65), Width = 300 };

                Label lParent = new Label { Text = "觸發條件 (父層選擇值)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 105) };
                _cboParentVals[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(0, 130), Width = 300 };
                
                Label lOpt = new Label { Text = "下拉選項內容 (每一行代表一個選項)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 170) };

                if (i == 0) {
                    lParent.Visible = false;
                    _cboParentVals[i].Visible = false;
                    lOpt.Location = new Point(0, 105);
                    pTopControls.Height = 135; // 第一層高度變矮，讓文字框更大
                }

                pTopControls.Controls.AddRange(new Control[] { lHeader, lCol, _cboCols[i], lParent, _cboParentVals[i], lOpt });

                _txtOptions[i] = new TextBox { 
                    Dock = DockStyle.Fill, 
                    Multiline = true, 
                    ScrollBars = ScrollBars.Both, 
                    WordWrap = false, 
                    Font = new Font("Microsoft JhengHei UI", 12F) 
                };

                // 自動隨視窗縮放寬度
                int captureIndex = i; // 閉包捕獲防呆
                pTopControls.Resize += (s, e) => {
                    if (captureIndex < 4) {
                        _cboCols[captureIndex].Width = pTopControls.Width - 5;
                        _cboParentVals[captureIndex].Width = pTopControls.Width - 5;
                    }
                };

                int colIndex = i;
                _cboCols[colIndex].SelectedIndexChanged += (s, e) => HandleColSelectionChanged(colIndex);
                if (i > 0) _cboParentVals[colIndex].SelectedIndexChanged += (s, e) => HandleParentValChanged(colIndex);

                pCol.Controls.Add(_txtOptions[i]);
                pCol.Controls.Add(pTopControls);

                tlpMain.Controls.Add(pCol, i, 0);
            }

            this.Controls.Add(tlpMain);

            // ================= 事件綁定 =================
            _btnExport.Click += BtnExport_Click;
            _btnImport.Click += BtnImport_Click;

            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) {
                foreach (var kvp in _dbMap) {
                    _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                }
            }
            
            _cboDb.SelectedIndexChanged += (s, e) => {
                _cboTable.Items.Clear();
                _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                ClearAllEditors();

                if (_cboDb.SelectedItem is ItemMap map && !string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                    foreach (var tbl in _dbMap[map.EnName].Tables) {
                        _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
                }
                if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
            };

            _cboTable.SelectedIndexChanged += (s, e) => {
                ClearAllEditors();
                if (_cboDb.SelectedItem is ItemMap dbMap && _cboTable.SelectedItem is ItemMap tbMap && !string.IsNullOrEmpty(dbMap.EnName) && !string.IsNullOrEmpty(tbMap.EnName)) {
                    var cols = DataManager.GetColumnNames(dbMap.EnName, tbMap.EnName);
                    foreach (var cbo in _cboCols) {
                        cbo.Items.Clear();
                        cbo.Items.Add("");
                        foreach (var c in cols) if (c != "Id" && c != "附件檔案" && c != "備註") cbo.Items.Add(c);
                    }
                }
            };
        }

        private void ClearAllEditors()
        {
            for (int i = 0; i < 4; i++) {
                if (_cboCols[i].Items.Count > 0) _cboCols[i].SelectedIndex = 0;
                if (i > 0) { _cboParentVals[i].Items.Clear(); _cboParentVals[i].Items.Add(""); }
                _txtOptions[i].Clear();
            }
        }

        private void HandleColSelectionChanged(int colIndex)
        {
            try {
                if (colIndex == 0)
                {
                    string colName = _cboCols[0].Text;
                    string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;
                    if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                        LoadOptionsToTextBox(tbName, colName, "", "", _txtOptions[0]);
                        UpdateChildParentVals(1, _txtOptions[0].Text);
                    }
                }
            } catch { }
        }

        private void HandleParentValChanged(int colIndex)
        {
            try {
                if (colIndex <= 0 || colIndex >= 4) return; // 絕對防越界
                
                string colName = _cboCols[colIndex].Text;
                string parentVal = _cboParentVals[colIndex].Text;
                string parentCol = _cboCols[colIndex - 1].Text;
                string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;

                if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                    LoadOptionsToTextBox(tbName, colName, parentCol, parentVal, _txtOptions[colIndex]);
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, _txtOptions[colIndex].Text);
                }
            } catch { }
        }

        private void UpdateChildParentVals(int childIndex, string parentOptionsText)
        {
            try {
                if (childIndex <= 0 || childIndex >= 4) return;
                
                var opts = parentOptionsText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                string currentVal = _cboParentVals[childIndex].Text;
                _cboParentVals[childIndex].Items.Clear();
                _cboParentVals[childIndex].Items.Add("");
                foreach (var o in opts) _cboParentVals[childIndex].Items.Add(o.Trim());

                if (!string.IsNullOrEmpty(currentVal) && _cboParentVals[childIndex].Items.Contains(currentVal))
                    _cboParentVals[childIndex].Text = currentVal;
                else
                    _cboParentVals[childIndex].SelectedIndex = 0;
            } catch { }
        }

        private void LoadOptionsToTextBox(string tableName, string colName, string parentColName, string parentVal, TextBox txt)
        {
            string opts = GetDropdownOptionsFromDB(tableName, colName, parentColName, parentVal);
            txt.Text = opts.Replace(",", Environment.NewLine);
        }

        private string GetDropdownOptionsFromDB(string tableName, string colName, string parentColName, string parentVal)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = "SELECT Options FROM DropdownConfigs WHERE TableName=@T AND ColName=@C AND IFNULL(ParentColName,'')=@PC AND IFNULL(ParentValue,'')=@PV";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@T", tableName);
                        cmd.Parameters.AddWithValue("@C", colName);
                        cmd.Parameters.AddWithValue("@PC", parentColName);
                        cmd.Parameters.AddWithValue("@PV", parentVal);
                        var res = cmd.ExecuteScalar();
                        return res != null ? res.ToString() : "";
                    }
                }
            } catch { return ""; }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_cboTable.SelectedItem == null || string.IsNullOrEmpty(((ItemMap)_cboTable.SelectedItem).EnName)) {
                MessageBox.Show("請先選擇資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        for (int i = 0; i < 4; i++) {
                            string colName = _cboCols[i].Text;
                            if (string.IsNullOrEmpty(colName)) continue;

                            string parentCol = i > 0 ? _cboCols[i-1].Text : "";
                            string parentVal = i > 0 ? _cboParentVals[i].Text : "";
                            
                            var optsArray = _txtOptions[i].Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToArray();
                            string optsStr = string.Join(",", optsArray);

                            string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) 
                                           VALUES (@T, @C, @PC, @PV, @Opt) 
                                           ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                            
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@T", tbName);
                                cmd.Parameters.AddWithValue("@C", colName);
                                cmd.Parameters.AddWithValue("@PC", parentCol);
                                cmd.Parameters.AddWithValue("@PV", parentVal);
                                cmd.Parameters.AddWithValue("@Opt", optsStr);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }

                MessageBox.Show("選項設定已儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDropdownConfigs(); 
                if (!string.IsNullOrEmpty(_txtOptions[0].Text)) UpdateChildParentVals(1, _txtOptions[0].Text);

            } catch (Exception ex) {
                MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================= 匯出與匯入 Excel =================
        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統下拉選單設定_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TableName AS [資料表名稱], ColName AS [欄位名稱], ParentColName AS [父層欄位], ParentValue AS [觸發條件], Options AS [選項內容(逗號分隔)] FROM DropdownConfigs", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("下拉選單設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！請直接在 Excel 中編輯後匯入。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string tb = ws.Cells[r, 1].Text.Trim();
                                        string col = ws.Cells[r, 2].Text.Trim();
                                        string pCol = ws.Cells[r, 3].Text.Trim();
                                        string pVal = ws.Cells[r, 4].Text.Trim();
                                        string opt = ws.Cells[r, 5].Text.Trim();

                                        if (string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) continue;

                                        string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) 
                                                       VALUES (@T, @C, @PC, @PV, @Opt) 
                                                       ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                                        
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@T", tb);
                                            cmd.Parameters.AddWithValue("@C", col);
                                            cmd.Parameters.AddWithValue("@PC", pCol);
                                            cmd.Parameters.AddWithValue("@PV", pVal);
                                            cmd.Parameters.AddWithValue("@Opt", opt);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("下拉選單設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDropdownConfigs();
                        _cboDb.SelectedIndex = 0; // 刷新畫面
                    } catch (Exception ex) {
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ================= 全域靜態快取供 TableLogic 使用 =================
        public static Dictionary<string, string[]> DropdownCache = new Dictionary<string, string[]>();

        public static void LoadDropdownConfigs()
        {
            DropdownCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM DropdownConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string tb = reader["TableName"].ToString();
                            string col = reader["ColName"].ToString();
                            string pCol = reader["ParentColName"].ToString();
                            string pVal = reader["ParentValue"].ToString();
                            string opts = reader["Options"].ToString();

                            string key = $"{tb}|{col}|{pCol}|{pVal}";
                            var arr = opts.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
                            DropdownCache[key] = arr;
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetOptions(string tableName, string colName, string parentColName = "", string parentValue = "")
        {
            string key = $"{tableName}|{colName}|{parentColName}|{parentValue}";
            if (DropdownCache.ContainsKey(key)) return DropdownCache[key];
            return null;
        }

        public static string[] GetAllOptionsForColumn(string tableName, string colName) 
        {
            HashSet<string> allOpts = new HashSet<string> { "" };
            foreach(var kvp in DropdownCache) {
                var parts = kvp.Key.Split('|');
                if(parts.Length == 4 && parts[0] == tableName && parts[1] == colName) {
                    foreach(var opt in kvp.Value) allOpts.Add(opt);
                }
            }
            return allOpts.ToArray();
        }
    }
}
