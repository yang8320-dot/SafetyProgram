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
        private Button _btnSave, _btnExport, _btnImport, _btnClearAll;
        
        // 🟢 快取已設定的項目，用於繪製深藍色字體
        private HashSet<string> _configuredDbs = new HashSet<string>();
        private HashSet<string> _configuredTables = new HashSet<string>();
        private HashSet<string> _configuredCols = new HashSet<string>();

        private bool _isRevertingDb = false;
        private bool _isRevertingCol = false;

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
                RefreshConfiguredCache();
                InitializeComponent();
                LoadDropdownConfigs();
            } catch (Exception ex) {
                MessageBox.Show($"初始化連動選單管理介面時發生嚴重錯誤：\n{ex.Message}\n{ex.StackTrace}", "系統崩潰防護", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 🟢 從資料庫撈取目前已有設定的表單與欄位，存入快取
        private void RefreshConfiguredCache()
        {
            _configuredDbs.Clear();
            _configuredTables.Clear();
            _configuredCols.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using(var cmd = new SQLiteCommand("SELECT DISTINCT TableName, ColName FROM DropdownConfigs", conn))
                    using(var reader = cmd.ExecuteReader()) {
                        while(reader.Read()) {
                            string tb = reader["TableName"].ToString();
                            string col = reader["ColName"].ToString();
                            _configuredTables.Add(tb);
                            _configuredCols.Add($"{tb}_{col}"); 
                        }
                    }
                }
                
                if (_dbMap != null) {
                    foreach(var kvp in _dbMap) {
                        foreach(var tb in kvp.Value.Tables.Keys) {
                            if (_configuredTables.Contains(tb)) {
                                _configuredDbs.Add(kvp.Key);
                            }
                        }
                    }
                }
            } catch {}
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
            
            _cboDb = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 30, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboDb.DrawItem += CboDb_DrawItem;

            Label lblTable = new Label { Text = "選擇資料表：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboTable = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 40, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboTable.DrawItem += CboTable_DrawItem;

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

            // 🟢 新增一鍵清除按鈕
            _btnClearAll = new Button { Text = "🗑️ 一鍵清除畫面上設定", Dock = DockStyle.Right, Width = 250, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnClearAll.Click += BtnClearAll_Click;

            Label lblHint = new Label { Text = "※ 已設定過下拉清單的項目，將以【深藍色】字體標示。\n※ 選項內容的排列順序，即為系統表單中下拉選單顯示的順序。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0, 5, 0, 0) };

            pnlBottom.Controls.Add(lblHint);
            
            // 加入 FlowLayoutPanel 控制按鈕排版 (靠右)
            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSave);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 }); // 間距
            flpBtnBottom.Controls.Add(_btnClearAll);
            
            pnlBottom.Controls.Add(flpBtnBottom);
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
                Panel pCol = new Panel { 
                    Dock = DockStyle.Fill, 
                    Margin = new Padding(3, 0, 3, 0), 
                    BackColor = Color.White, 
                    Padding = new Padding(15) 
                };
                pCol.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pCol.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                Panel pTopControls = new Panel { Dock = DockStyle.Top, Height = 195, BackColor = Color.White };

                Label lHeader = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Location = new Point(0, 0) };
                
                Label lCol = new Label { Text = "綁定資料表欄位：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 40) };
                _cboCols[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(0, 65), Width = 300, DrawMode = DrawMode.OwnerDrawFixed };
                
                int currentIndex = i; // 閉包捕獲
                _cboCols[i].DrawItem += (s, e) => CboCols_DrawItem(s, e, currentIndex);

                Label lParent = new Label { Text = "觸發條件 (父層選擇值)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 105) };
                _cboParentVals[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(0, 130), Width = 300 };
                
                Label lOpt = new Label { Text = "下拉選項內容 (每一行代表一個選項)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(0, 170) };

                if (i == 0) {
                    lParent.Visible = false;
                    _cboParentVals[i].Visible = false;
                    lOpt.Location = new Point(0, 105);
                    pTopControls.Height = 135; 
                }

                pTopControls.Controls.AddRange(new Control[] { lHeader, lCol, _cboCols[i], lParent, _cboParentVals[i], lOpt });

                _txtOptions[i] = new TextBox { 
                    Dock = DockStyle.Fill, 
                    Multiline = true, 
                    ScrollBars = ScrollBars.Both, 
                    WordWrap = false, 
                    Font = new Font("Microsoft JhengHei UI", 12F) 
                };

                int captureIndex = i; 
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
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
        }

        // ================= 🟢 自訂繪製邏輯 (深藍色高亮) =================
        private void CboDb_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var item = _cboDb.Items[e.Index] as ItemMap;
            bool isConfig = item != null && !string.IsNullOrEmpty(item.EnName) && _configuredDbs.Contains(item.EnName);
            DrawComboBoxItem(_cboDb, e, isConfig);
        }

        private void CboTable_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var item = _cboTable.Items[e.Index] as ItemMap;
            bool isConfig = item != null && !string.IsNullOrEmpty(item.EnName) && _configuredTables.Contains(item.EnName);
            DrawComboBoxItem(_cboTable, e, isConfig);
        }

        private void CboCols_DrawItem(object sender, DrawItemEventArgs e, int colIndex)
        {
            if (e.Index < 0) return;
            string colName = _cboCols[colIndex].Items[e.Index].ToString();
            string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName ?? "";
            bool isConfig = !string.IsNullOrEmpty(colName) && _configuredCols.Contains($"{tbName}_{colName}");
            DrawComboBoxItem(_cboCols[colIndex], e, isConfig);
        }

        private void DrawComboBoxItem(ComboBox cbo, DrawItemEventArgs e, bool isConfigured)
        {
            e.DrawBackground();
            string text = cbo.Items[e.Index].ToString();
            
            Brush textBrush = Brushes.Black;
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) {
                textBrush = Brushes.White;
            } else if (isConfigured) {
                textBrush = Brushes.DarkBlue; // 🟢 已設定項目呈現深藍色
            }
            
            e.Graphics.DrawString(text, e.Font, textBrush, e.Bounds);
            e.DrawFocusRectangle();
        }

        // ================= 🟢 一鍵清除畫面上設定 =================
        private void BtnClearAll_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("確定要清空畫面上所有的選項與連動設定嗎？\n(注意：尚未按下儲存前，資料庫內的設定並不會被刪除。)", "確認清除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                ClearAllEditors();
            }
        }

        // ================= 🟢 隱藏選單密碼防護 =================
        private bool VerifyHiddenMenuPassword(string menuName)
        {
            using (Form p = new Form())
            {
                p.Width = 460; 
                p.Height = 220;
                p.Text = "隱藏選單安全驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.White;

                Label lbl = new Label() { Left = 30, Top = 30, Text = $"請輸入【{menuName}】的解鎖密碼以繼續設定：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 250, Left = 30, Top = 70, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認驗證", DialogResult = DialogResult.OK, Left = 160, Top = 120, Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog(this) == DialogResult.OK)
                {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName) return true;
                    
                    MessageBox.Show($"【{menuName}】密碼錯誤！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false; 
            }
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isRevertingDb) return;

            var selectedDb = _cboDb.SelectedItem as ItemMap;
            if (selectedDb != null && selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB"))
            {
                string menuName = "";
                if (selectedDb.EnName == "Menu1DB") menuName = "選單1";
                if (selectedDb.EnName == "Menu2DB") menuName = "選單2";
                if (selectedDb.EnName == "Menu3DB") menuName = "選單3";
                if (selectedDb.EnName == "Menu4DB") menuName = "選單4";

                if (!string.IsNullOrEmpty(menuName))
                {
                    if (!VerifyHiddenMenuPassword(menuName))
                    {
                        _isRevertingDb = true;
                        _cboDb.SelectedIndex = 0; // 退回空白
                        _isRevertingDb = false;
                        return;
                    }
                }
            }

            _cboTable.Items.Clear();
            _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
            ClearAllEditors();

            if (selectedDb != null && !string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                foreach (var tbl in _dbMap[selectedDb.EnName].Tables) {
                    _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            }
            if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
        }

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ClearAllEditors();
            if (_cboDb.SelectedItem is ItemMap dbMap && _cboTable.SelectedItem is ItemMap tbMap && !string.IsNullOrEmpty(dbMap.EnName) && !string.IsNullOrEmpty(tbMap.EnName)) {
                var cols = DataManager.GetColumnNames(dbMap.EnName, tbMap.EnName);
                foreach (var cbo in _cboCols) {
                    cbo.Items.Clear();
                    cbo.Items.Add("");
                    foreach (var c in cols) if (c != "Id" && c != "附件檔案" && c != "備註") cbo.Items.Add(c);
                }
            }
        }

        private void ClearAllEditors()
        {
            _isRevertingCol = true;
            for (int i = 0; i < 4; i++) {
                if (_cboCols[i].Items.Count > 0) _cboCols[i].SelectedIndex = 0;
                if (i > 0) { _cboParentVals[i].Items.Clear(); _cboParentVals[i].Items.Add(""); }
                _txtOptions[i].Clear();
            }
            _isRevertingCol = false;
        }

        private void HandleColSelectionChanged(int colIndex)
        {
            if (_isRevertingCol) return;

            string selectedCol = _cboCols[colIndex].Text;
            
            // 🟢 重複欄位防呆機制
            if (!string.IsNullOrEmpty(selectedCol))
            {
                for (int i = 0; i < 4; i++)
                {
                    if (i != colIndex && _cboCols[i].Text == selectedCol)
                    {
                        MessageBox.Show("此欄位已在其他層級被設定，為防止系統錯亂，請勿重複選擇！", "重複選擇防呆", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isRevertingCol = true;
                        _cboCols[colIndex].SelectedIndex = 0; // 強制退回空白
                        _isRevertingCol = false;
                        return;
                    }
                }
            }

            try {
                if (colIndex == 0)
                {
                    string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;
                    if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(selectedCol)) {
                        LoadOptionsToTextBox(tbName, selectedCol, "", "", _txtOptions[0]);
                        UpdateChildParentVals(1, _txtOptions[0].Text);
                    } else {
                        _txtOptions[0].Clear();
                        UpdateChildParentVals(1, "");
                    }
                }
            } catch { }
        }

        private void HandleParentValChanged(int colIndex)
        {
            try {
                if (colIndex <= 0 || colIndex >= 4) return;
                
                string colName = _cboCols[colIndex].Text;
                string parentVal = _cboParentVals[colIndex].Text;
                string parentCol = _cboCols[colIndex - 1].Text;
                string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;

                if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                    LoadOptionsToTextBox(tbName, colName, parentCol, parentVal, _txtOptions[colIndex]);
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, _txtOptions[colIndex].Text);
                } else {
                    _txtOptions[colIndex].Clear();
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, "");
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
                
                RefreshConfiguredCache(); 
                _cboDb.Invalidate(); 
                _cboTable.Invalidate(); 
                foreach(var c in _cboCols) c.Invalidate(); 

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
                        
                        RefreshConfiguredCache();
                        LoadDropdownConfigs();
                        _cboDb.SelectedIndex = 0; 
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
