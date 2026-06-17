/// FILE: Safety_System/settings/App_DropdownManager.TabMulti.cs ///
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
        // ================= Tab 2: 組合文字 (複選) 控制項 =================
        private ComboBox _cboDbMulti, _cboTableMulti, _cboColMulti;
        private DataGridView _dgvOptionsMulti; 
        private Button _btnSaveMulti, _btnDelMulti;
        private FlowLayoutPanel _flpMultiConfigured;
        private Panel _selectedMultiItemPanel = null; 
        private bool _isRevertingMultiCol = false;

        private void BuildTabMulti(TabPage page)
        {
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 95, BackColor = Color.White, Padding = new Padding(20) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSaveMulti = new Button { Text = "💾 儲存組合文字設定", Width = 230, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSaveMulti.Click += BtnSaveMulti_Click;

            _btnDelMulti = new Button { Text = "🗑️ 刪除此欄位設定", Width = 230, Height = 50, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnDelMulti.Click += BtnDelMulti_Click;

            Button btnClearMultiDb = new Button { Text = "💣 清除所有資料庫設定", Width = 260, Height = 50, BackColor = Color.Crimson, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnClearMultiDb.Click += BtnClearMultiDb_Click;

            Label lblHintMulti = new Label { Text = "※ 已設定組合文字的項目，以【紅色粗體】標示。\n※ 拖曳列首(把手)可排序。資料表多選時將以並排圖示顯示。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0) };
            
            pnlBottom.Controls.Add(lblHintMulti);

            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSaveMulti);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 });
            flpBtnBottom.Controls.Add(_btnDelMulti);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 });
            flpBtnBottom.Controls.Add(btnClearMultiDb);
            
            pnlBottom.Controls.Add(flpBtnBottom);
            page.Controls.Add(pnlBottom);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            Label lblTitle = new Label { Text = "☑️ 組合文字 (複選) 設定區 (支援自訂圖示)", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkCyan, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };

            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

            Label l1 = new Label { Text = "選擇資料庫：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboDbMulti = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 20, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboDbMulti.DrawItem += CboDbMulti_DrawItem;
            
            _cboDbMulti.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) { foreach (var kvp in _dbMap) _cboDbMulti.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName }); }

            Label l2 = new Label { Text = "選擇資料表：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboTableMulti = new ComboBox { Width = 260, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 20, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboTableMulti.DrawItem += CboTableMulti_DrawItem;

            Label l3 = new Label { Text = "指定目標欄位：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboColMulti = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 30, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboColMulti.DrawItem += CboColMulti_DrawItem;

            Button btnExportMulti = new Button { Text = "📤 匯出 Excel", Size = new Size(150, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 15, 0) };
            btnExportMulti.Click += BtnExportMulti_Click;

            Button btnImportMulti = new Button { Text = "📥 匯入 Excel", Size = new Size(150, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            btnImportMulti.Click += BtnImportMulti_Click;

            flpControls.Controls.AddRange(new Control[] { l1, _cboDbMulti, l2, _cboTableMulti, l3, _cboColMulti, btnExportMulti, btnImportMulti });
            flpTopMain.Controls.Add(lblTitle);
            flpTopMain.Controls.Add(flpControls);
            pnlTop.Controls.Add(flpTopMain);
            page.Controls.Add(pnlTop);

            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            Panel pnlLeftBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(15) };
            pnlLeftBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlLeftBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            FlowLayoutPanel flpMultiOptHeader = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 10), WrapContents = false };
            Label l4 = new Label { Text = "自訂選項內容：", Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 5, 2, 0) };
            
            Button btnExtractMulti = new Button { Text = "導入", Size = new Size(60, 30), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0) };
            btnExtractMulti.FlatAppearance.BorderSize = 0;
            btnExtractMulti.Click += (s, e) => ExtractDataFromDB(0, true);
            
            Button btnMultiUp = new Button { Text = "↑", Size = new Size(35, 30), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(5, 0, 0, 0) };
            btnMultiUp.FlatAppearance.BorderSize = 0;
            btnMultiUp.Click += (s, e) => MoveRowUp(_dgvOptionsMulti);

            Button btnMultiDown = new Button { Text = "↓", Size = new Size(35, 30), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(5, 0, 0, 0) };
            btnMultiDown.FlatAppearance.BorderSize = 0;
            btnMultiDown.Click += (s, e) => MoveRowDown(_dgvOptionsMulti);

            flpMultiOptHeader.Controls.Add(l4);
            flpMultiOptHeader.Controls.Add(btnExtractMulti);
            flpMultiOptHeader.Controls.Add(btnMultiUp);
            flpMultiOptHeader.Controls.Add(btnMultiDown);

            _dgvOptionsMulti = new DataGridView { 
                Dock = DockStyle.Fill, 
                AllowUserToAddRows = true, 
                AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                RowHeadersVisible = true, 
                RowHeadersWidth = 35,
                EditMode = DataGridViewEditMode.EditOnEnter, 
                BackgroundColor = Color.WhiteSmoke,
                RowTemplate = { Height = 45 }
            };
            
            _dgvOptionsMulti.AllowDrop = true;
            _dgvOptionsMulti.MouseDown += DgvOptions_MouseDown;
            _dgvOptionsMulti.MouseMove += DgvOptions_MouseMove;
            _dgvOptionsMulti.DragOver += DgvOptions_DragOver;
            _dgvOptionsMulti.DragDrop += DgvOptions_DragDrop;

            _dgvOptionsMulti.Columns.Add(new DataGridViewTextBoxColumn { Name = "Text", HeaderText = "選項文字", FillWeight = 50 });
            _dgvOptionsMulti.Columns.Add(new DataGridViewImageColumn { Name = "Icon", HeaderText = "圖示預覽", ImageLayout = DataGridViewImageCellLayout.Zoom, FillWeight = 20 });
            _dgvOptionsMulti.Columns.Add(new DataGridViewButtonColumn { Name = "Upload", HeaderText = "上傳", Text = "上傳", UseColumnTextForButtonValue = true, FillWeight = 15 });
            _dgvOptionsMulti.Columns.Add(new DataGridViewButtonColumn { Name = "Clear", HeaderText = "清除", Text = "清除", UseColumnTextForButtonValue = true, FillWeight = 15 });

            _dgvOptionsMulti.CellContentClick += (s, e) => DgvOptions_CellContentClick(s, e, 0, true);
            
            pnlLeftBorder.Controls.Add(_dgvOptionsMulti);
            pnlLeftBorder.Controls.Add(flpMultiOptHeader); 

            Panel pnlRightBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(15) };
            pnlRightBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlRightBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Label l5 = new Label { Text = "已設定之組合文字清單：", Dock = DockStyle.Top, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Margin = new Padding(0, 0, 0, 10), Height = 30 };
            _flpMultiConfigured = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5) };
            
            pnlRightBorder.Controls.Add(_flpMultiConfigured);
            pnlRightBorder.Controls.Add(l5);

            tlpMain.Controls.Add(pnlLeftBorder, 0, 0);
            tlpMain.Controls.Add(pnlRightBorder, 1, 0);

            page.Controls.Add(tlpMain);
            tlpMain.BringToFront();

            _cboDbMulti.SelectedIndexChanged += (s, e) => {
                _cboTableMulti.Items.Clear(); _cboTableMulti.Items.Add(new ItemMap { EnName = "", ChName = "" }); _cboColMulti.Items.Clear(); _dgvOptionsMulti.Rows.Clear();
                var db = _cboDbMulti.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tbl in _dbMap[db.EnName].Tables) _cboTableMulti.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            };
            
            _cboTableMulti.SelectedIndexChanged += (s, e) => {
                _cboColMulti.Items.Clear(); _dgvOptionsMulti.Rows.Clear();
                var db = _cboDbMulti.SelectedItem as ItemMap; var tb = _cboTableMulti.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    var cols = GetColumnsSafe(db.EnName, tb.EnName).Where(c => c != "Id" && c != "附件檔案");
                    foreach (var c in cols) _cboColMulti.Items.Add(c);
                }
            };

            _cboColMulti.SelectedIndexChanged += CboColMulti_SelectedIndexChanged_Event;
        }

        private void CboDbMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboDbMulti.Items[e.Index] as ItemMap;
            bool isConfig = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName)) {
                foreach(var k in MultiSelectCache.Keys) {
                    string tb = k.Split('|')[0];
                    if (_dbMap.ContainsKey(item.EnName) && _dbMap[item.EnName].Tables.ContainsKey(tb)) { isConfig = true; break; }
                }
            }
            DrawComboBoxItemMulti(_cboDbMulti, e, isConfig, false);
        }

        private void CboTableMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboTableMulti.Items[e.Index] as ItemMap;
            bool isConfig = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName)) {
                foreach(var k in MultiSelectCache.Keys) {
                    if (k.StartsWith(item.EnName + "|")) { isConfig = true; break; }
                }
            }
            DrawComboBoxItemMulti(_cboTableMulti, e, isConfig, false);
        }

        private void CboColMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            string col = _cboColMulti.Items[e.Index].ToString();
            var db = _cboDbMulti.SelectedItem as ItemMap;
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            bool isConfig = false;
            bool isOtherConfig = false; 

            if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName) && !string.IsNullOrEmpty(col)) {
                string key = $"{tb.EnName}|{col}";
                if (MultiSelectCache.ContainsKey(key)) {
                    isConfig = true;
                } else if (IsColumnInDropdownCache(tb.EnName, col) || ReferenceCache.ContainsKey($"{db.EnName}|{tb.EnName}|{col}")) {
                    isOtherConfig = true;
                }
            }
            DrawComboBoxItemMulti(_cboColMulti, e, isConfig, isOtherConfig);
        }

        private void DrawComboBoxItemMulti(ComboBox cbo, DrawItemEventArgs e, bool isConfigured, bool isOtherConfigured) {
            if (e.Index < 0) return;
            
            Brush bgBrush = Brushes.White;
            Brush textBrush = Brushes.Black;
            Font currentFont = e.Font;

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) {
                bgBrush = SystemBrushes.Highlight;
                textBrush = Brushes.White;
                if (isConfigured || isOtherConfigured) currentFont = new Font(e.Font, FontStyle.Bold);
            } else {
                if (isOtherConfigured) {
                    bgBrush = Brushes.LightGray; 
                    textBrush = Brushes.DimGray;
                    currentFont = new Font(e.Font, FontStyle.Bold);
                } else if (isConfigured) {
                    textBrush = Brushes.Crimson; 
                    currentFont = new Font(e.Font, FontStyle.Bold);
                }
            }
            
            e.Graphics.FillRectangle(bgBrush, e.Bounds);
            e.Graphics.DrawString(cbo.Items[e.Index].ToString(), currentFont, textBrush, new RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height));
            e.DrawFocusRectangle();
        }

        private void CboColMulti_SelectedIndexChanged_Event(object sender, EventArgs e) {
            if (_isRevertingMultiCol) return;

            _dgvOptionsMulti.Rows.Clear();
            var db = _cboDbMulti.SelectedItem as ItemMap;
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            if (db != null && tb != null && _cboColMulti.SelectedItem != null) {
                string colName = _cboColMulti.SelectedItem.ToString();
                
                string conflict = CheckColumnConflict(db.EnName, tb.EnName, colName, "TabMulti");
                if (conflict != null) {
                    MessageBox.Show($"此欄位【{colName}】已在 {conflict} 中設定過！\n為避免系統判斷異常，同一欄位不可重複設定為不同型態。\n\n請先前往該分頁刪除設定後再試。", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _isRevertingMultiCol = true;
                    _cboColMulti.SelectedIndex = -1; 
                    _isRevertingMultiCol = false;
                    return;
                }

                LoadOptionsToGrid(tb.EnName, colName, "", "", _dgvOptionsMulti, true);
            }
        }

        private void BtnSaveMulti_Click(object sender, EventArgs e)
        {
            var db = _cboDbMulti.SelectedItem as ItemMap;
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            if (db == null || tb == null || _cboColMulti.SelectedItem == null || _dgvOptionsMulti.Rows.Count <= 1) {
                MessageBox.Show("請確認資料庫、資料表、欄位與選項內容皆已填寫！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = db.EnName;
            string tblName = tb.EnName;
            string colName = _cboColMulti.SelectedItem.ToString();
            
            string conflict = CheckColumnConflict(dbName, tblName, colName, "TabMulti");
            if (conflict != null) {
                MessageBox.Show($"欄位【{colName}】已在 {conflict} 中設定過！無法儲存！", "儲存攔截", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            List<string> optList = new List<string>();
            foreach(DataGridViewRow r in _dgvOptionsMulti.Rows) {
                if (r.IsNewRow) continue;
                string text = r.Cells["Text"].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(text)) continue;
                
                string b64 = r.Cells["Icon"].Tag?.ToString() ?? "";
                if (!string.IsNullOrEmpty(b64)) {
                    optList.Add($"{text}[ICON:{b64}]");
                } else {
                    optList.Add(text);
                }
            }
            string optsStr = string.Join(",", optList);

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"INSERT INTO MultiSelectConfigs (TableName, ColName, Options) VALUES (@T, @C, @Opt) ON CONFLICT(TableName, ColName) DO UPDATE SET Options=@Opt";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@T", tblName);
                        cmd.Parameters.AddWithValue("@C", colName);
                        cmd.Parameters.AddWithValue("@Opt", optsStr);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("組合文字設定儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadMultiSelectConfigs();
                RefreshMultiConfiguredList();
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message, "錯誤"); }
        }

        private void BtnDelMulti_Click(object sender, EventArgs e)
        {
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            if (tb == null || _cboColMulti.SelectedItem == null) return;

            if (MessageBox.Show("確定要刪除此欄位的組合文字設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs WHERE TableName=@T AND ColName=@C", conn)) {
                            cmd.Parameters.AddWithValue("@T", tb.EnName);
                            cmd.Parameters.AddWithValue("@C", _cboColMulti.SelectedItem.ToString());
                            cmd.ExecuteNonQuery();
                        }
                    }
                    _dgvOptionsMulti.Rows.Clear();
                    LoadMultiSelectConfigs();
                    RefreshMultiConfiguredList();
                    MessageBox.Show("刪除成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message, "錯誤"); }
            }
        }

        private void BtnClearMultiDb_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("【極度危險】\n您確定要永久刪除資料庫中「所有」的組合文字(複選)設定嗎？\n此操作無法復原！", "永久刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
            {
                string authPrompt = "清除設定資料需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return;

                try 
                {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs", conn)) 
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    
                    MessageBox.Show("資料庫內的所有組合文字設定已全部清空！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    _dgvOptionsMulti.Rows.Clear();
                    LoadMultiSelectConfigs();
                    RefreshMultiConfiguredList();
                } 
                catch (Exception ex) 
                {
                    MessageBox.Show($"清除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnExportMulti_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統組合文字(複選)設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                        {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TableName AS [資料表名稱], ColName AS [欄位名稱], Options AS [選項內容(純文字)] FROM MultiSelectConfigs", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        foreach(DataRow row in dt.Rows) {
                            string rawOpts = row["選項內容(純文字)"].ToString();
                            var parts = rawOpts.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
                            for(int i=0; i<parts.Length; i++) {
                                int idx = parts[i].IndexOf("[ICON:");
                                if(idx >= 0) parts[i] = parts[i].Substring(0, idx);
                            }
                            row["選項內容(純文字)"] = string.Join(",", parts);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("組合文字設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("組合文字設定匯出成功！請直接在 Excel 中編輯後匯入。\n注意：Excel不支援夾帶圖示，匯出檔案僅包含純文字。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImportMulti_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                            {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) 
                                {
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) 
                                    {
                                        string tb = ws.Cells[r, 1].Text.Trim();
                                        string col = ws.Cells[r, 2].Text.Trim();
                                        string opt = ws.Cells[r, 3].Text.Trim();

                                        if (string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) continue;

                                        string sql = @"INSERT INTO MultiSelectConfigs (TableName, ColName, Options) 
                                                       VALUES (@T, @C, @Opt) 
                                                       ON CONFLICT(TableName, ColName) DO UPDATE SET Options=@Opt";
                                        
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) 
                                        {
                                            cmd.Parameters.AddWithValue("@T", tb);
                                            cmd.Parameters.AddWithValue("@C", col);
                                            cmd.Parameters.AddWithValue("@Opt", opt);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        
                        MessageBox.Show("組合文字設定已批次匯入並【自動存檔覆寫】成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        _dgvOptionsMulti.Rows.Clear();
                        LoadMultiSelectConfigs();
                        RefreshMultiConfiguredList();
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void RefreshMultiConfiguredList()
        {
            if (_flpMultiConfigured == null) return;
            _flpMultiConfigured.Controls.Clear();
            _selectedMultiItemPanel = null; 
            
            if (MultiSelectCache.Count == 0) {
                _flpMultiConfigured.Controls.Add(new Label { Text = "尚無任何設定。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                return;
            }

            foreach (var kvp in MultiSelectCache) {
                string[] parts = kvp.Key.Split('|');
                if (parts.Length != 2) continue;

                string tbName = parts[0];
                string colName = parts[1];
                string chTbName = tbName;
                string dbName = ""; 

                foreach (var db in _dbMap) {
                    if (db.Value.Tables.ContainsKey(tbName)) {
                        chTbName = db.Value.Tables[tbName];
                        dbName = db.Key;
                        break;
                    }
                }

                Panel pItem = new Panel { Width = 650, Height = 45, BackColor = Color.AliceBlue, Margin = new Padding(5), Cursor = Cursors.Hand };
                
                Button btnDel = new Button { 
                    Text = "❌", 
                    Location = new Point(10, 7), 
                    Size = new Size(30, 30), 
                    FlatStyle = FlatStyle.Flat, 
                    ForeColor = Color.IndianRed, 
                    BackColor = Color.Transparent,
                    Cursor = Cursors.Hand,
                    Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold)
                };
                btnDel.FlatAppearance.BorderSize = 0;
                btnDel.FlatAppearance.MouseOverBackColor = Color.MistyRose;
                
                btnDel.Click += (s, e) => {
                    if (MessageBox.Show($"確定要刪除【{chTbName} - {colName}】的組合文字設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        try {
                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs WHERE TableName=@T AND ColName=@C", conn)) {
                                    cmd.Parameters.AddWithValue("@T", tbName);
                                    cmd.Parameters.AddWithValue("@C", colName);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            _dgvOptionsMulti.Rows.Clear();
                            LoadMultiSelectConfigs();
                            RefreshMultiConfiguredList();
                            MessageBox.Show("刪除成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message, "錯誤"); }
                    }
                };

                Label lName = new Label { 
                    Text = $"表：{chTbName}   ➡️   欄位：[{colName}]", 
                    Location = new Point(50, 12), 
                    AutoSize = true, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = Color.DarkSlateBlue,
                    Cursor = Cursors.Hand
                };

                Action selectAction = () => {
                    if (_selectedMultiItemPanel != null && _selectedMultiItemPanel != pItem) {
                        _selectedMultiItemPanel.BackColor = Color.AliceBlue; 
                    }
                    pItem.BackColor = Color.LightSkyBlue; 
                    _selectedMultiItemPanel = pItem;

                    foreach (ItemMap item in _cboDbMulti.Items) {
                        if (item.EnName == dbName) { _cboDbMulti.SelectedItem = item; break; }
                    }
                    foreach (ItemMap item in _cboTableMulti.Items) {
                        if (item.EnName == tbName) { _cboTableMulti.SelectedItem = item; break; }
                    }
                    if (_cboColMulti.Items.Contains(colName)) {
                        _cboColMulti.SelectedItem = colName;
                    }
                };

                pItem.Click += (s, e) => selectAction();
                lName.Click += (s, e) => selectAction();

                pItem.Controls.Add(btnDel);
                pItem.Controls.Add(lName);
                _flpMultiConfigured.Controls.Add(pItem);
            }
            
            _flpMultiConfigured.Resize -= FlpMultiConfigured_Resize; 
            _flpMultiConfigured.Resize += FlpMultiConfigured_Resize;
        }

        private void FlpMultiConfigured_Resize(object sender, EventArgs e)
        {
            foreach (Control ctrl in _flpMultiConfigured.Controls)
            {
                if (ctrl is Panel pnl)
                {
                    pnl.Width = _flpMultiConfigured.ClientSize.Width - 20;
                }
            }
        }

        public static void LoadMultiSelectConfigs()
        {
            MultiSelectCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmdChk = new SQLiteCommand("CREATE TABLE IF NOT EXISTS [MultiSelectConfigs] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TableName TEXT, ColName TEXT, Options TEXT, UNIQUE(TableName, ColName));", conn)) { cmdChk.ExecuteNonQuery(); }

                    using (var cmd = new SQLiteCommand("SELECT TableName, ColName, Options FROM MultiSelectConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string key = $"{reader["TableName"]}|{reader["ColName"]}";
                            string opts = reader["Options"].ToString();
                            
                            var arr = opts.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            List<DropdownItemDef> defs = new List<DropdownItemDef>();
                            
                            foreach(var opt in arr) {
                                string text = opt.Trim(); string b64 = "";
                                int idx = opt.IndexOf("[ICON:");
                                if (idx >= 0 && opt.EndsWith("]")) {
                                    text = opt.Substring(0, idx).Trim();
                                    b64 = opt.Substring(idx + 6, opt.Length - idx - 7);
                                }
                                if(!string.IsNullOrEmpty(text)) {
                                    defs.Add(new DropdownItemDef { Text = text, IconBase64 = b64 });
                                }
                            }
                            MultiSelectCache[key] = defs;
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetMultiSelectOptions(string tableName, string colName)
        {
            string key = $"{tableName}|{colName}";
            if (MultiSelectCache.ContainsKey(key)) return MultiSelectCache[key].Select(d => d.Text).ToArray();
            return null;
        }
    }
}
