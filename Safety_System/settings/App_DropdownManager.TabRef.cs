/// FILE: Safety_System/settings/App_DropdownManager.TabRef.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public partial class App_DropdownManager
    {
        // ================= Tab 3: 跨表參照下拉選單控制項 =================
        private ComboBox _cboRefTargetDb, _cboRefTargetTb, _cboRefTargetCol;
        private ComboBox _cboRefSourceDb, _cboRefSourceTb, _cboRefSourceCol;
        private Button _btnSaveRef, _btnDelRef, _btnExportRef, _btnImportRef;
        private FlowLayoutPanel _flpRefConfigured;
        private Panel _selectedRefItemPanel = null;
        private bool _isRevertingRefCol = false;

        public static void LoadReferenceConfigs()
        {
            ReferenceCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT TargetDb, TargetTb, TargetCol, SourceDb, SourceTb, SourceCol FROM ReferenceDropdownConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string key = $"{reader["TargetDb"]}|{reader["TargetTb"]}|{reader["TargetCol"]}";
                            ReferenceCache[key] = new ReferenceDef {
                                SourceDb = reader["SourceDb"].ToString(),
                                SourceTb = reader["SourceTb"].ToString(),
                                SourceCol = reader["SourceCol"].ToString()
                            };
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetReferenceOptions(string targetDb, string targetTb, string targetCol)
        {
            string key = $"{targetDb}|{targetTb}|{targetCol}";
            if (!ReferenceCache.ContainsKey(key)) return null;

            var def = ReferenceCache[key];
            HashSet<string> uniqueValues = new HashSet<string> { "" }; 
            
            try {
                DataTable dt = DataManager.GetTableData(def.SourceDb, def.SourceTb, "", "", "");
                if (dt != null && dt.Columns.Contains(def.SourceCol)) {
                    foreach (DataRow r in dt.Rows) {
                        if (r.RowState == DataRowState.Deleted) continue;
                        string val = r[def.SourceCol]?.ToString().Trim() ?? "";
                        if (!string.IsNullOrEmpty(val)) {
                            uniqueValues.Add(val);
                        }
                    }
                }
            } catch { }

            return uniqueValues.ToArray();
        }

        private void BuildTabReference(TabPage page)
        {
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 95, BackColor = Color.White, Padding = new Padding(20) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSaveRef = new Button { Text = "💾 儲存跨表參照設定", Width = 230, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSaveRef.Click += BtnSaveRef_Click;

            _btnDelRef = new Button { Text = "🗑️ 刪除此欄位設定", Width = 230, Height = 50, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnDelRef.Click += BtnDelRef_Click;

            Label lblHintRef = new Label { Text = "※ 此功能可讓下拉選單「動態」讀取另一個資料表的特定欄位內容。\n※ 例如：讓下拉選單直接呈現所有已經建檔的【單號】或【組合文字】，無須手動輸入。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0) };
            
            pnlBottom.Controls.Add(lblHintRef);

            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSaveRef);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 });
            flpBtnBottom.Controls.Add(_btnDelRef);
            
            pnlBottom.Controls.Add(flpBtnBottom);
            page.Controls.Add(pnlBottom);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            
            Label lblTitle = new Label { Text = "🔗 跨表參照下拉選單設定區", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.SaddleBrown, AutoSize = true, Margin = new Padding(0, 0, 20, 15) };
            
            _btnExportRef = new Button { Text = "📤 匯出 Excel", Size = new Size(150, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 15, 0) };
            _btnExportRef.Click += BtnExportRef_Click;

            _btnImportRef = new Button { Text = "📥 匯入 Excel", Size = new Size(150, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnImportRef.Click += BtnImportRef_Click;

            flpTopMain.Controls.Add(lblTitle);
            flpTopMain.Controls.Add(_btnExportRef);
            flpTopMain.Controls.Add(_btnImportRef);
            
            pnlTop.Controls.Add(flpTopMain);
            page.Controls.Add(pnlTop);

            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // 左側：設定區
            Panel pnlLeftBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(25) };
            pnlLeftBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlLeftBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            FlowLayoutPanel flpSettings = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            // 1. 目標欄位設定
            Label lTargetTitle = new Label { Text = "🎯 目標欄位 (我要在哪裡產生下拉選單？)", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            _cboRefTargetDb = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboRefTargetTb = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboRefTargetCol = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };

            flpSettings.Controls.Add(lTargetTitle);
            flpSettings.Controls.Add(CreateRefRow("目標資料庫：", _cboRefTargetDb));
            flpSettings.Controls.Add(CreateRefRow("目標資料表：", _cboRefTargetTb));
            flpSettings.Controls.Add(CreateRefRow("目標欄位：", _cboRefTargetCol));

            flpSettings.Controls.Add(new Panel { Width = 500, Height = 2, BackColor = Color.LightGray, Margin = new Padding(0, 20, 0, 20) });

            // 2. 來源欄位設定
            Label lSourceTitle = new Label { Text = "📦 來源資料 (下拉選單的選項要從哪裡抓？)", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkGreen, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            _cboRefSourceDb = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboRefSourceTb = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboRefSourceCol = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };

            flpSettings.Controls.Add(lSourceTitle);
            flpSettings.Controls.Add(CreateRefRow("來源資料庫：", _cboRefSourceDb));
            flpSettings.Controls.Add(CreateRefRow("來源資料表：", _cboRefSourceTb));
            flpSettings.Controls.Add(CreateRefRow("來源欄位：", _cboRefSourceCol));

            pnlLeftBorder.Controls.Add(flpSettings);

            // 右側：已設定清單區
            Panel pnlRightBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(15) };
            pnlRightBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlRightBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Label lListTitle = new Label { Text = "已設定之跨表參照清單：", Dock = DockStyle.Top, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Margin = new Padding(0, 0, 0, 10), Height = 30 };
            _flpRefConfigured = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5) };
            
            pnlRightBorder.Controls.Add(_flpRefConfigured);
            pnlRightBorder.Controls.Add(lListTitle);

            tlpMain.Controls.Add(pnlLeftBorder, 0, 0);
            tlpMain.Controls.Add(pnlRightBorder, 1, 0);

            page.Controls.Add(tlpMain);
            tlpMain.BringToFront();

            BindRefComboBoxEvents();
        }

       private Panel CreateRefRow(string labelText, ComboBox cbo)
        {
            Panel p = new Panel { Width = 600, Height = 45, Margin = new Padding(0, 0, 0, 5) };
            Label l = new Label { Text = labelText, Location = new Point(0, 5), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            cbo.Location = new Point(140, 2);
            p.Controls.Add(l); p.Controls.Add(cbo);
            return p;
        }

        private void BindRefComboBoxEvents()
        {
            _cboRefTargetDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            _cboRefSourceDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) { 
                foreach (var kvp in _dbMap) {
                    _cboRefTargetDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                    _cboRefSourceDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                }
            }

            // Target Events
            _cboRefTargetDb.SelectedIndexChanged += (s, e) => {
                _cboRefTargetTb.Items.Clear(); _cboRefTargetTb.Items.Add(new ItemMap { EnName = "", ChName = "" }); _cboRefTargetCol.Items.Clear();
                var db = _cboRefTargetDb.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tbl in _dbMap[db.EnName].Tables) _cboRefTargetTb.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            };
            
            _cboRefTargetTb.SelectedIndexChanged += (s, e) => {
                _cboRefTargetCol.Items.Clear();
                var db = _cboRefTargetDb.SelectedItem as ItemMap; var tb = _cboRefTargetTb.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    var cols = GetColumnsSafe(db.EnName, tb.EnName).Where(c => c != "Id" && c != "附件檔案");
                    foreach (var c in cols) _cboRefTargetCol.Items.Add(c);
                }
            };

            _cboRefTargetCol.SelectedIndexChanged += (s, e) => {
                if (_isRevertingRefCol) return;
                var db = _cboRefTargetDb.SelectedItem as ItemMap; var tb = _cboRefTargetTb.SelectedItem as ItemMap;
                if (db != null && tb != null && _cboRefTargetCol.SelectedItem != null) {
                    string colName = _cboRefTargetCol.SelectedItem.ToString();
                    
                    string conflict = CheckColumnConflict(db.EnName, tb.EnName, colName, "TabRef");
                    if (conflict != null) {
                        MessageBox.Show($"此欄位【{colName}】已在 {conflict} 中設定過！\n為避免系統判斷異常，同一欄位不可重複設定為不同型態。\n\n請先前往該分頁刪除設定後再試。", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isRevertingRefCol = true;
                        _cboRefTargetCol.SelectedIndex = -1; 
                        _isRevertingRefCol = false;
                        return;
                    }
                }
            };

            // Source Events
            _cboRefSourceDb.SelectedIndexChanged += (s, e) => {
                _cboRefSourceTb.Items.Clear(); _cboRefSourceTb.Items.Add(new ItemMap { EnName = "", ChName = "" }); _cboRefSourceCol.Items.Clear();
                var db = _cboRefSourceDb.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tbl in _dbMap[db.EnName].Tables) _cboRefSourceTb.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            };
            
            _cboRefSourceTb.SelectedIndexChanged += (s, e) => {
                _cboRefSourceCol.Items.Clear();
                var db = _cboRefSourceDb.SelectedItem as ItemMap; var tb = _cboRefSourceTb.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    var cols = GetColumnsSafe(db.EnName, tb.EnName).Where(c => c != "Id" && c != "附件檔案");
                    foreach (var c in cols) _cboRefSourceCol.Items.Add(c);
                }
            };
        }

        private void BtnSaveRef_Click(object sender, EventArgs e)
        {
            var tDb = _cboRefTargetDb.SelectedItem as ItemMap; var tTb = _cboRefTargetTb.SelectedItem as ItemMap;
            var sDb = _cboRefSourceDb.SelectedItem as ItemMap; var sTb = _cboRefSourceTb.SelectedItem as ItemMap;

            if (tDb == null || tTb == null || _cboRefTargetCol.SelectedItem == null ||
                sDb == null || sTb == null || _cboRefSourceCol.SelectedItem == null ||
                string.IsNullOrEmpty(tDb.EnName) || string.IsNullOrEmpty(tTb.EnName) ||
                string.IsNullOrEmpty(sDb.EnName) || string.IsNullOrEmpty(sTb.EnName)) 
            {
                MessageBox.Show("請確認目標欄位與來源欄位皆已完整選擇！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                return;
            }

            string tCol = _cboRefTargetCol.SelectedItem.ToString();
            string sCol = _cboRefSourceCol.SelectedItem.ToString();
            
            string conflict = CheckColumnConflict(tDb.EnName, tTb.EnName, tCol, "TabRef");
            if (conflict != null) {
                MessageBox.Show($"欄位【{tCol}】已在 {conflict} 中設定過！無法儲存！", "儲存攔截", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"INSERT INTO ReferenceDropdownConfigs (TargetDb, TargetTb, TargetCol, SourceDb, SourceTb, SourceCol) 
                                   VALUES (@TD, @TT, @TC, @SD, @ST, @SC) 
                                   ON CONFLICT(TargetDb, TargetTb, TargetCol) DO UPDATE SET SourceDb=@SD, SourceTb=@ST, SourceCol=@SC";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@TD", tDb.EnName);
                        cmd.Parameters.AddWithValue("@TT", tTb.EnName);
                        cmd.Parameters.AddWithValue("@TC", tCol);
                        cmd.Parameters.AddWithValue("@SD", sDb.EnName);
                        cmd.Parameters.AddWithValue("@ST", sTb.EnName);
                        cmd.Parameters.AddWithValue("@SC", sCol);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("跨表參照設定儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadReferenceConfigs();
                RefreshRefConfiguredList();
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message, "錯誤"); }
        }

        private void BtnDelRef_Click(object sender, EventArgs e)
        {
            var tDb = _cboRefTargetDb.SelectedItem as ItemMap; 
            var tTb = _cboRefTargetTb.SelectedItem as ItemMap;
            if (tDb == null || tTb == null || _cboRefTargetCol.SelectedItem == null) return;

            if (MessageBox.Show("確定要刪除此欄位的跨表參照設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM ReferenceDropdownConfigs WHERE TargetDb=@TD AND TargetTb=@TT AND TargetCol=@TC", conn)) {
                            cmd.Parameters.AddWithValue("@TD", tDb.EnName);
                            cmd.Parameters.AddWithValue("@TT", tTb.EnName);
                            cmd.Parameters.AddWithValue("@TC", _cboRefTargetCol.SelectedItem.ToString());
                            cmd.ExecuteNonQuery();
                        }
                    }
                    LoadReferenceConfigs();
                    RefreshRefConfiguredList();
                    MessageBox.Show("刪除成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message, "錯誤"); }
            }
        }

        private void RefreshRefConfiguredList()
        {
            if (_flpRefConfigured == null) return;
            _flpRefConfigured.Controls.Clear();
            _selectedRefItemPanel = null; 
            
            if (ReferenceCache.Count == 0) {
                _flpRefConfigured.Controls.Add(new Label { Text = "尚無任何設定。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                return;
            }

            foreach (var kvp in ReferenceCache) {
                string[] parts = kvp.Key.Split('|');
                if (parts.Length != 3) continue;

                string tDbName = parts[0];
                string tTbName = parts[1];
                string tColName = parts[2];
                var srcDef = kvp.Value;

                string chTTbName = tTbName;
                string chSTbName = srcDef.SourceTb;

                if (_dbMap.ContainsKey(tDbName) && _dbMap[tDbName].Tables.ContainsKey(tTbName)) chTTbName = _dbMap[tDbName].Tables[tTbName];
                if (_dbMap.ContainsKey(srcDef.SourceDb) && _dbMap[srcDef.SourceDb].Tables.ContainsKey(srcDef.SourceTb)) chSTbName = _dbMap[srcDef.SourceDb].Tables[srcDef.SourceTb];

                Panel pItem = new Panel { Width = 700, Height = 65, BackColor = Color.Honeydew, Margin = new Padding(5), Cursor = Cursors.Hand };
                pItem.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pItem.ClientRectangle, Color.LightGreen, ButtonBorderStyle.Solid);
                
                Button btnDel = new Button { 
                    Text = "❌", Location = new Point(10, 15), Size = new Size(30, 30), 
                    FlatStyle = FlatStyle.Flat, ForeColor = Color.IndianRed, BackColor = Color.Transparent, Cursor = Cursors.Hand
                };
                btnDel.FlatAppearance.BorderSize = 0;
                
                btnDel.Click += (s, e) => {
                    if (MessageBox.Show($"確定要刪除【{chTTbName} - {tColName}】的跨表設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        try {
                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var cmd = new SQLiteCommand("DELETE FROM ReferenceDropdownConfigs WHERE TargetDb=@TD AND TargetTb=@TT AND TargetCol=@TC", conn)) {
                                    cmd.Parameters.AddWithValue("@TD", tDbName);
                                    cmd.Parameters.AddWithValue("@TT", tTbName);
                                    cmd.Parameters.AddWithValue("@TC", tColName);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            LoadReferenceConfigs();
                            RefreshRefConfiguredList();
                        } catch { }
                    }
                };

                Label lTarget = new Label { Text = $"🎯 目標：{chTTbName} [{tColName}]", Location = new Point(50, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkBlue, Cursor = Cursors.Hand };
                Label lSource = new Label { Text = $"📦 來源：{chSTbName} [{srcDef.SourceCol}]", Location = new Point(50, 35), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkGreen, Cursor = Cursors.Hand };

                Action selectAction = () => {
                    if (_selectedRefItemPanel != null && _selectedRefItemPanel != pItem) {
                        _selectedRefItemPanel.BackColor = Color.Honeydew; 
                    }
                    pItem.BackColor = Color.PaleGreen; 
                    _selectedRefItemPanel = pItem;

                    foreach (ItemMap item in _cboRefTargetDb.Items) { if (item.EnName == tDbName) { _cboRefTargetDb.SelectedItem = item; break; } }
                    foreach (ItemMap item in _cboRefTargetTb.Items) { if (item.EnName == tTbName) { _cboRefTargetTb.SelectedItem = item; break; } }
                    if (_cboRefTargetCol.Items.Contains(tColName)) _cboRefTargetCol.SelectedItem = tColName;

                    foreach (ItemMap item in _cboRefSourceDb.Items) { if (item.EnName == srcDef.SourceDb) { _cboRefSourceDb.SelectedItem = item; break; } }
                    foreach (ItemMap item in _cboRefSourceTb.Items) { if (item.EnName == srcDef.SourceTb) { _cboRefSourceTb.SelectedItem = item; break; } }
                    if (_cboRefSourceCol.Items.Contains(srcDef.SourceCol)) _cboRefSourceCol.SelectedItem = srcDef.SourceCol;
                };

                pItem.Click += (s, e) => selectAction();
                lTarget.Click += (s, e) => selectAction();
                lSource.Click += (s, e) => selectAction();

                pItem.Controls.Add(btnDel);
                pItem.Controls.Add(lTarget);
                pItem.Controls.Add(lSource);
                _flpRefConfigured.Controls.Add(pItem);
            }
            
            _flpRefConfigured.Resize -= FlpRefConfigured_Resize; 
            _flpRefConfigured.Resize += FlpRefConfigured_Resize;
        }

        private void FlpRefConfigured_Resize(object sender, EventArgs e)
        {
            foreach (Control ctrl in _flpRefConfigured.Controls)
            {
                if (ctrl is Panel pnl)
                {
                    pnl.Width = _flpRefConfigured.ClientSize.Width - 20;
                }
            }
        }

        // ========================================================
        // 🟢 新增：跨表參照的匯出/匯入 Excel 功能
        // ========================================================
        private void BtnExportRef_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "跨表參照下拉選單設定_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TargetDb AS [目標資料庫], TargetTb AS [目標資料表], TargetCol AS [目標欄位], SourceDb AS [來源資料庫], SourceTb AS [來源資料表], SourceCol AS [來源欄位] FROM ReferenceDropdownConfigs", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }
                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("跨表參照設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("跨表參照設定匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { 
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }

        private void BtnImportRef_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的跨表參照設定檔" })
            {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string tDb = ws.Cells[r, 1].Text.Trim();
                                        string tTb = ws.Cells[r, 2].Text.Trim();
                                        string tCol = ws.Cells[r, 3].Text.Trim();
                                        string sDb = ws.Cells[r, 4].Text.Trim();
                                        string sTb = ws.Cells[r, 5].Text.Trim();
                                        string sCol = ws.Cells[r, 6].Text.Trim();

                                        if (string.IsNullOrEmpty(tDb) || string.IsNullOrEmpty(tTb) || string.IsNullOrEmpty(tCol) ||
                                            string.IsNullOrEmpty(sDb) || string.IsNullOrEmpty(sTb) || string.IsNullOrEmpty(sCol)) continue;

                                        string sql = @"INSERT INTO ReferenceDropdownConfigs (TargetDb, TargetTb, TargetCol, SourceDb, SourceTb, SourceCol) 
                                                       VALUES (@TD, @TT, @TC, @SD, @ST, @SC) 
                                                       ON CONFLICT(TargetDb, TargetTb, TargetCol) DO UPDATE SET SourceDb=@SD, SourceTb=@ST, SourceCol=@SC";
                                        
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@TD", tDb);
                                            cmd.Parameters.AddWithValue("@TT", tTb);
                                            cmd.Parameters.AddWithValue("@TC", tCol);
                                            cmd.Parameters.AddWithValue("@SD", sDb);
                                            cmd.Parameters.AddWithValue("@ST", sTb);
                                            cmd.Parameters.AddWithValue("@SC", sCol);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        
                        MessageBox.Show("跨表參照設定已批次匯入並【自動存檔覆寫】成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadReferenceConfigs();
                        RefreshRefConfiguredList();
                    } 
                    catch (Exception ex) { 
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }
    }
}
