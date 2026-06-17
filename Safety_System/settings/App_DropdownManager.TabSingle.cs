/// FILE: Safety_System/settings/App_DropdownManager.TabSingle.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public partial class App_DropdownManager
    {
        // ================= Tab 1: 單選連動下拉選單控制項 =================
        private ComboBox _cboDb, _cboTable;
        private DataGridView[] _dgvOptions;
        private ComboBox[] _cboCols;
        private ComboBox[] _cboParentVals;
        private Button _btnSave, _btnExport, _btnImport, _btnClearAll, _btnClearDb;
        
        private Dictionary<string, bool> _configuredDbs = new Dictionary<string, bool>();
        private Dictionary<string, bool> _configuredTables = new Dictionary<string, bool>();
        private Dictionary<string, bool> _configuredCols = new Dictionary<string, bool>();
        private Dictionary<string, int> _configuredColLevels = new Dictionary<string, int>();

        private bool _isRevertingDb = false;
        private bool _isRevertingCol = false;

        private void RefreshConfiguredCache()
        {
            _configuredDbs.Clear();
            _configuredTables.Clear();
            _configuredCols.Clear();
            _configuredColLevels.Clear();

            Dictionary<string, string> parentMap = new Dictionary<string, string>();
            HashSet<string> allConfigured = new HashSet<string>();

            try 
            {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                {
                    conn.Open();
                    using(var cmd = new SQLiteCommand("SELECT TableName, ColName, ParentColName, Options FROM DropdownConfigs", conn))
                    using(var reader = cmd.ExecuteReader()) 
                    {
                        while(reader.Read()) 
                        {
                            string opts = reader["Options"].ToString();
                            if (string.IsNullOrWhiteSpace(opts)) continue; 

                            string tb = reader["TableName"].ToString();
                            string col = reader["ColName"].ToString();
                            string pCol = reader["ParentColName"].ToString();

                            string colKey = $"{tb}_{col}";
                            allConfigured.Add(colKey);
                            _configuredTables[tb] = false; 

                            if (!string.IsNullOrEmpty(pCol)) 
                            {
                                string pColKey = $"{tb}_{pCol}";
                                parentMap[colKey] = pColKey;
                                allConfigured.Add(pColKey);
                            }
                        }
                    }
                }
                
                foreach (string colKey in allConfigured) 
                {
                    bool isLinked = parentMap.ContainsKey(colKey) || parentMap.Values.Contains(colKey);
                    _configuredCols[colKey] = isLinked;

                    string tb = colKey.Split('_')[0];
                    if (isLinked) 
                    {
                        _configuredTables[tb] = true;
                        
                        int level = 1;
                        string curr = colKey;
                        int safeguard = 0;
                        while (parentMap.ContainsKey(curr)) 
                        {
                            level++;
                            curr = parentMap[curr];
                            if (safeguard++ > 10) break;
                        }
                        _configuredColLevels[colKey] = level;
                    }
                }

                if (_dbMap != null) 
                {
                    foreach(var kvp in _dbMap) 
                    {
                        string dbName = kvp.Key;
                        foreach(var tb in kvp.Value.Tables.Keys) 
                        {
                            if (_configuredTables.ContainsKey(tb)) 
                            {
                                if (!_configuredDbs.ContainsKey(dbName)) _configuredDbs[dbName] = false;
                                if (_configuredTables[tb]) _configuredDbs[dbName] = true;
                            }
                        }
                    }
                }
            } 
            catch { }
        }

        private void BuildTabSingle(TabPage page)
        {
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 95, BackColor = Color.White, Padding = new Padding(20) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSave = new Button { Text = "💾 儲存並套用當前設定", Width = 260, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSave.Click += BtnSave_Click;

            _btnClearAll = new Button { Text = "🗑️ 一鍵清除畫面上設定", Width = 260, Height = 50, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnClearAll.Click += BtnClearAll_Click;

            _btnClearDb = new Button { Text = "💣 清除所有資料庫設定", Width = 260, Height = 50, BackColor = Color.Crimson, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnClearDb.Click += BtnClearDb_Click;

            Label lblHint = new Label { Text = "※ 已設定過且「僅單層獨立」的項目，以【亮藍色粗體】標示。\n※ 拖曳表格最左側列首(把手)或點擊上下按鈕可調整選項順序。\n※ 點擊儲存格即可直接打字修改，無需連點兩下。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0) };

            pnlBottom.Controls.Add(lblHint);
            
            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSave);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 }); 
            flpBtnBottom.Controls.Add(_btnClearAll);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 }); 
            flpBtnBottom.Controls.Add(_btnClearDb);
            
            pnlBottom.Controls.Add(flpBtnBottom);
            page.Controls.Add(pnlBottom); 

            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            Label lblTitle = new Label { Text = "🔧 單選下拉與多層連動設定 (支援自訂圖示)", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
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
            page.Controls.Add(pnlTop); 

            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            for(int i = 0; i < 4; i++) tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            _cboCols = new ComboBox[4];
            _cboParentVals = new ComboBox[4];
            _dgvOptions = new DataGridView[4];

            string[] headers = { "第一層 (主選項)", "第二層 (依第一層連動)", "第三層 (依第二層連動)", "第四層 (依第三層連動)" };

            for (int i = 0; i < 4; i++)
            {
                Panel pColBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(3, 0, 3, 0), BackColor = Color.White, Padding = new Padding(15) };
                pColBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pColBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                TableLayoutPanel pColInner = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7 };
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

                Label lHeader = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0,0,0,15) };
                
                Label lCol = new Label { Text = "綁定資料表欄位：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,0,0,5) };
                _cboCols[i] = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0,0,0,15), DrawMode = DrawMode.OwnerDrawFixed };
                int currentIndex = i; 
                _cboCols[i].DrawItem += (s, e) => CboCols_DrawItem(s, e, currentIndex);

                Label lParent = new Label { Text = "觸發條件 (父層選擇值)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,0,0,5) };
                _cboParentVals[i] = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0,0,0,15) };
                
                FlowLayoutPanel flpOptHeader = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0,0,0,5), WrapContents = false };
                Label lOpt = new Label { Text = "自訂選項內容：", Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,5,2,0) };
                
                Button btnExtract = new Button { Text = "導入", Size = new Size(60, 30), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0) };
                btnExtract.FlatAppearance.BorderSize = 0;
                btnExtract.Click += (s, e) => ExtractDataFromDB(currentIndex, false);

                Button btnUp = new Button { Text = "↑", Size = new Size(35, 30), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(5, 0, 0, 0) };
                btnUp.FlatAppearance.BorderSize = 0;
                btnUp.Click += (s, e) => MoveRowUp(_dgvOptions[currentIndex]);

                Button btnDown = new Button { Text = "↓", Size = new Size(35, 30), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(5, 0, 0, 0) };
                btnDown.FlatAppearance.BorderSize = 0;
                btnDown.Click += (s, e) => MoveRowDown(_dgvOptions[currentIndex]);

                flpOptHeader.Controls.Add(lOpt);
                flpOptHeader.Controls.Add(btnExtract);
                flpOptHeader.Controls.Add(btnUp);
                flpOptHeader.Controls.Add(btnDown);

                _dgvOptions[i] = new DataGridView { 
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
                
                _dgvOptions[i].AllowDrop = true;
                _dgvOptions[i].MouseDown += DgvOptions_MouseDown;
                _dgvOptions[i].MouseMove += DgvOptions_MouseMove;
                _dgvOptions[i].DragOver += DgvOptions_DragOver;
                _dgvOptions[i].DragDrop += DgvOptions_DragDrop;

                _dgvOptions[i].Columns.Add(new DataGridViewTextBoxColumn { Name = "Text", HeaderText = "選項文字", FillWeight = 50 });
                _dgvOptions[i].Columns.Add(new DataGridViewImageColumn { Name = "Icon", HeaderText = "預覽", ImageLayout = DataGridViewImageCellLayout.Zoom, FillWeight = 20 });
                _dgvOptions[i].Columns.Add(new DataGridViewButtonColumn { Name = "Upload", HeaderText = "上傳", Text = "上傳", UseColumnTextForButtonValue = true, FillWeight = 15 });
                _dgvOptions[i].Columns.Add(new DataGridViewButtonColumn { Name = "Clear", HeaderText = "清除", Text = "清除", UseColumnTextForButtonValue = true, FillWeight = 15 });

                int dgvIndex = i;
                _dgvOptions[i].CellContentClick += (s, e) => DgvOptions_CellContentClick(s, e, dgvIndex, false);

                if (i == 0) {
                    lParent.Visible = false;
                    _cboParentVals[i].Visible = false;
                }

                pColInner.Controls.Add(lHeader, 0, 0);
                pColInner.Controls.Add(lCol, 0, 1);
                pColInner.Controls.Add(_cboCols[i], 0, 2);
                pColInner.Controls.Add(lParent, 0, 3);
                pColInner.Controls.Add(_cboParentVals[i], 0, 4);
                pColInner.Controls.Add(flpOptHeader, 0, 5);
                pColInner.Controls.Add(_dgvOptions[i], 0, 6);

                int colIndex = i;
                _cboCols[colIndex].SelectedIndexChanged += (s, e) => HandleColSelectionChanged(colIndex);
                if (i > 0) _cboParentVals[colIndex].SelectedIndexChanged += (s, e) => HandleParentValChanged(colIndex);

                pColBorder.Controls.Add(pColInner);
                tlpMain.Controls.Add(pColBorder, i, 0);
            }

            page.Controls.Add(tlpMain); 
            tlpMain.BringToFront();     

            _btnExport.Click += BtnExport_Click;
            _btnImport.Click += BtnImport_Click;

            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) 
            {
                foreach (var kvp in _dbMap) _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            }
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
        }

        private void CboDb_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboDb.Items[e.Index] as ItemMap;
            bool isConfig = false;
            bool isLinked = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName) && _configuredDbs.ContainsKey(item.EnName)) {
                isConfig = true; isLinked = _configuredDbs[item.EnName];
            }
            DrawComboBoxItem(_cboDb, e, isConfig, isLinked, false);
        }

        private void CboTable_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboTable.Items[e.Index] as ItemMap;
            bool isConfig = false;
            bool isLinked = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName) && _configuredTables.ContainsKey(item.EnName)) {
                isConfig = true; isLinked = _configuredTables[item.EnName];
            }
            DrawComboBoxItem(_cboTable, e, isConfig, isLinked, false);
        }

        private void CboCols_DrawItem(object sender, DrawItemEventArgs e, int colIndex) {
            if (e.Index < 0) return;
            string colName = _cboCols[colIndex].Items[e.Index].ToString();
            string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName ?? "";
            string dbName = ((ItemMap)_cboDb.SelectedItem)?.EnName ?? "";
            
            bool isConfig = false;
            bool isLinked = false;
            bool isOtherConfig = false; 
            int level = 0;
            
            string key = $"{tbName}_{colName}";
            if (!string.IsNullOrEmpty(colName) && _configuredCols.ContainsKey(key)) {
                isConfig = true; 
                isLinked = _configuredCols[key];
                if (_configuredColLevels.ContainsKey(key)) {
                    level = _configuredColLevels[key];
                }
            } else if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                if (MultiSelectCache.ContainsKey($"{tbName}|{colName}") || ReferenceCache.ContainsKey($"{dbName}|{tbName}|{colName}")) {
                    isOtherConfig = true;
                }
            }
            DrawComboBoxItem(_cboCols[colIndex], e, isConfig, isLinked, isOtherConfig, level);
        }

        private void DrawComboBoxItem(ComboBox cbo, DrawItemEventArgs e, bool isConfigured, bool isLinked, bool isOtherConfigured, int level = 0) {
            if (e.Index < 0) return;

            Brush bgBrush = Brushes.White;
            Brush textBrush = Brushes.Black;
            Font currentFont = e.Font;
            
            string displayText = cbo.Items[e.Index].ToString();
            if (isLinked && level > 0) {
                displayText += $" (第{level}層)";
            }

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
                    textBrush = isLinked ? Brushes.Crimson : Brushes.DodgerBlue; 
                    currentFont = new Font(e.Font, FontStyle.Bold);         
                }
            }
            
            e.Graphics.FillRectangle(bgBrush, e.Bounds);
            e.Graphics.DrawString(displayText, currentFont, textBrush, new RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height));
            e.DrawFocusRectangle();
        }

        private void BtnClearAll_Click(object sender, EventArgs e) {
            if (MessageBox.Show("確定要清空畫面上所有的選項與連動設定嗎？\n(注意：尚未按下儲存前，資料庫內的設定並不會被刪除。)", "確認清除畫面", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) {
                ClearAllEditors();
            }
        }

        private void BtnClearDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show("【極度危險】\n您確定要永久刪除資料庫中「所有」的單選下拉與連動設定嗎？\n此操作無法復原！", "永久刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM DropdownConfigs", conn)) { cmd.ExecuteNonQuery(); }
                    }
                    MessageBox.Show("資料庫設定已全部清空！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ClearAllEditors();
                    RefreshConfiguredCache();
                    LoadDropdownConfigs();
                    
                    _cboDb.SelectedIndex = 0; 
                    _cboDb.Invalidate();
                    _cboTable.Invalidate();
                    foreach(var c in _cboCols) c.Invalidate();
                } catch (Exception ex) {
                    MessageBox.Show($"清除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool VerifyHiddenMenuPassword(string menuName) {
            using (Form p = new Form()) {
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

                if (p.ShowDialog(this) == DialogResult.OK) {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName) return true;
                    
                    MessageBox.Show($"【{menuName}】密碼錯誤！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false; 
            }
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e) {
            if (_isRevertingDb) return;
            var selectedDb = _cboDb.SelectedItem as ItemMap;
            if (selectedDb != null && selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB")) {
                string menuName = "";
                if (selectedDb.EnName == "Menu1DB") menuName = "選單1";
                if (selectedDb.EnName == "Menu2DB") menuName = "選單2";
                if (selectedDb.EnName == "Menu3DB") menuName = "選單3";
                if (selectedDb.EnName == "Menu4DB") menuName = "選單4";

                if (!string.IsNullOrEmpty(menuName)) {
                    if (!VerifyHiddenMenuPassword(menuName)) {
                        _isRevertingDb = true;
                        _cboDb.SelectedIndex = 0; 
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

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e) {
            ClearAllEditors();
            if (_cboDb.SelectedItem is ItemMap dbMap && _cboTable.SelectedItem is ItemMap tbMap && !string.IsNullOrEmpty(dbMap.EnName) && !string.IsNullOrEmpty(tbMap.EnName)) {
                var cols = GetColumnsSafe(dbMap.EnName, tbMap.EnName);
                foreach (var cbo in _cboCols) {
                    cbo.Items.Clear();
                    cbo.Items.Add("");
                    foreach (var c in cols) if (c != "Id" && c != "附件檔案" && c != "備註") cbo.Items.Add(c);
                }
            }
        }

        private void ClearAllEditors() {
            _isRevertingCol = true;
            for (int i = 0; i < 4; i++) {
                if (_cboCols[i].Items.Count > 0) _cboCols[i].SelectedIndex = 0;
                if (i > 0) { _cboParentVals[i].Items.Clear(); _cboParentVals[i].Items.Add(""); }
                _dgvOptions[i].Rows.Clear();
            }
            _isRevertingCol = false;
        }

        private void HandleColSelectionChanged(int colIndex) {
            if (_isRevertingCol) return;
            string selectedCol = _cboCols[colIndex].Text;
            string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName ?? "";
            string dbName = ((ItemMap)_cboDb.SelectedItem)?.EnName ?? "";
            
            if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(selectedCol)) {
                
                string conflict = CheckColumnConflict(dbName, tbName, selectedCol, "TabSingle");
                if (conflict != null) {
                    MessageBox.Show($"此欄位【{selectedCol}】已在 {conflict} 中設定過！\n為避免系統判斷異常，同一欄位不可重複設定為不同型態。\n\n請先前往該分頁刪除設定後再試。", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _isRevertingCol = true;
                    _cboCols[colIndex].SelectedIndex = 0; 
                    _isRevertingCol = false;
                    return;
                }

                for (int i = 0; i < 4; i++) {
                    if (i != colIndex && _cboCols[i].Text == selectedCol) {
                        MessageBox.Show("此欄位已在本頁面其他層級被設定，為防止系統錯亂，請勿重複選擇！", "重複選擇防呆", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isRevertingCol = true;
                        _cboCols[colIndex].SelectedIndex = 0; 
                        _isRevertingCol = false;
                        return;
                    }
                }
            }

            try {
                if (colIndex == 0) {
                    if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(selectedCol)) {
                        LoadOptionsToGrid(tbName, selectedCol, "", "", _dgvOptions[0], false);
                        UpdateChildParentVals(1, _dgvOptions[0]);
                    } else {
                        _dgvOptions[0].Rows.Clear();
                        UpdateChildParentVals(1, _dgvOptions[0]);
                    }
                }
            } catch { }
        }

        private void HandleParentValChanged(int colIndex) {
            try {
                if (colIndex <= 0 || colIndex >= 4) return;
                
                string colName = _cboCols[colIndex].Text;
                string parentVal = _cboParentVals[colIndex].Text;
                string parentCol = _cboCols[colIndex - 1].Text;
                string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;

                if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                    LoadOptionsToGrid(tbName, colName, parentCol, parentVal, _dgvOptions[colIndex], false);
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, _dgvOptions[colIndex]);
                } else {
                    _dgvOptions[colIndex].Rows.Clear();
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, _dgvOptions[colIndex]);
                }
            } catch { }
        }

        private void UpdateChildParentVals(int childIndex, DataGridView dgvParent) {
            try {
                if (childIndex <= 0 || childIndex >= 4) return;
                
                string currentVal = _cboParentVals[childIndex].Text;
                _cboParentVals[childIndex].Items.Clear();
                _cboParentVals[childIndex].Items.Add("");
                
                foreach(DataGridViewRow r in dgvParent.Rows) {
                    if (r.IsNewRow) continue;
                    string txt = r.Cells["Text"].Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(txt)) {
                        _cboParentVals[childIndex].Items.Add(txt.Trim());
                    }
                }

                if (!string.IsNullOrEmpty(currentVal) && _cboParentVals[childIndex].Items.Contains(currentVal))
                    _cboParentVals[childIndex].Text = currentVal;
                else
                    _cboParentVals[childIndex].SelectedIndex = 0;
            } catch { }
        }

        private void LoadOptionsToGrid(string tableName, string colName, string parentColName, string parentVal, DataGridView dgv, bool isMulti) {
            dgv.Rows.Clear();
            
            string optsStr = "";
            if (isMulti) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("SELECT Options FROM MultiSelectConfigs WHERE TableName=@T AND ColName=@C", conn)) {
                            cmd.Parameters.AddWithValue("@T", tableName);
                            cmd.Parameters.AddWithValue("@C", colName);
                            var res = cmd.ExecuteScalar();
                            if (res != null) optsStr = res.ToString();
                        }
                    }
                } catch { }
            } else {
                optsStr = GetDropdownOptionsFromDB(tableName, colName, parentColName, parentVal);
            }
            
            var parts = optsStr.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
            foreach(var p in parts) {
                string text = p.Trim(); 
                string b64 = "";
                int idx = p.IndexOf("[ICON:");
                if (idx >= 0 && p.EndsWith("]")) {
                    text = p.Substring(0, idx).Trim();
                    b64 = p.Substring(idx + 6, p.Length - idx - 7);
                }
                
                if (string.IsNullOrWhiteSpace(text)) continue;

                int rIdx = dgv.Rows.Add();
                dgv.Rows[rIdx].Cells["Text"].Value = text;
                if (!string.IsNullOrEmpty(b64)) {
                    try {
                        byte[] bytes = Convert.FromBase64String(b64);
                        using(MemoryStream ms = new MemoryStream(bytes)) {
                            dgv.Rows[rIdx].Cells["Icon"].Value = Image.FromStream(ms);
                        }
                        dgv.Rows[rIdx].Cells["Icon"].Tag = b64;
                    } catch { }
                }
            }
        }

        private string GetDropdownOptionsFromDB(string tableName, string colName, string parentColName, string parentVal) {
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

        private void ExtractDataFromDB(int colIndex, bool isMulti)
        {
            string db = isMulti ? ((ItemMap)_cboDbMulti.SelectedItem)?.EnName : ((ItemMap)_cboDb.SelectedItem)?.EnName;
            string tb = isMulti ? ((ItemMap)_cboTableMulti.SelectedItem)?.EnName : ((ItemMap)_cboTable.SelectedItem)?.EnName;
            string col = isMulti ? _cboColMulti.SelectedItem?.ToString() : _cboCols[colIndex].Text;

            if (string.IsNullOrEmpty(db) || string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) {
                MessageBox.Show("請先正確選擇資料庫、資料表與欄位！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try {
                DataTable dt = DataManager.GetTableData(db, tb, "", "", "");
                if (dt == null || dt.Rows.Count == 0) {
                    MessageBox.Show("此資料表目前尚無資料可供導入。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                HashSet<string> uniqueVals = new HashSet<string>();
                foreach(DataRow r in dt.Rows) {
                    if (dt.Columns.Contains(col)) {
                        string v = r[col]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(v)) {
                            if (isMulti) {
                                var parts = v.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach(var p in parts) {
                                    string cleanP = p.Trim();
                                    if (!string.IsNullOrEmpty(cleanP)) uniqueVals.Add(cleanP);
                                }
                            } else {
                                uniqueVals.Add(v);
                            }
                        }
                    }
                }

                if (uniqueVals.Count > 0) {
                    int addedCount = 0;
                    DataGridView dgv = isMulti ? _dgvOptionsMulti : _dgvOptions[colIndex];
                    List<string> existing = new List<string>();
                    
                    foreach(DataGridViewRow r in dgv.Rows) {
                        if (!r.IsNewRow && r.Cells["Text"].Value != null) existing.Add(r.Cells["Text"].Value.ToString().Trim());
                    }
                    
                    foreach(var v in uniqueVals) {
                        if (!existing.Contains(v)) {
                            int rIdx = dgv.Rows.Add();
                            dgv.Rows[rIdx].Cells["Text"].Value = v;
                            addedCount++;
                        }
                    }
                    
                    MessageBox.Show($"導入完成！共新增 {addedCount} 筆不重複資料至設定中。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("導入完成，但在該欄位沒有發現有效資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch (Exception ex) {
                MessageBox.Show("導入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            if (_cboTable.SelectedItem == null || string.IsNullOrEmpty(((ItemMap)_cboTable.SelectedItem).EnName)) {
                MessageBox.Show("請先選擇資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
            string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;

            for (int i = 0; i < 4; i++) {
                string colName = _cboCols[i].Text;
                if (!string.IsNullOrEmpty(colName)) {
                    string conflict = CheckColumnConflict(dbName, tbName, colName, "TabSingle");
                    if (conflict != null) {
                        MessageBox.Show($"欄位【{colName}】已在 {conflict} 中設定過！\n為避免系統錯亂，不可同時設定為多種選單模式，無法儲存！", "儲存攔截", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        for (int i = 0; i < 4; i++) {
                            string colName = _cboCols[i].Text;
                            if (string.IsNullOrEmpty(colName)) continue;

                            string parentCol = i > 0 ? _cboCols[i-1].Text : "";
                            string parentVal = i > 0 ? _cboParentVals[i].Text : "";
                            
                            List<string> optList = new List<string>();
                            foreach(DataGridViewRow r in _dgvOptions[i].Rows) {
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

                            if (string.IsNullOrEmpty(optsStr))
                            {
                                string sqlDel = @"DELETE FROM DropdownConfigs 
                                                  WHERE TableName=@T AND ColName=@C AND IFNULL(ParentColName,'')=@PC AND IFNULL(ParentValue,'')=@PV";
                                using (var cmd = new SQLiteCommand(sqlDel, conn, trans)) {
                                    cmd.Parameters.AddWithValue("@T", tbName); cmd.Parameters.AddWithValue("@C", colName);
                                    cmd.Parameters.AddWithValue("@PC", parentCol); cmd.Parameters.AddWithValue("@PV", parentVal);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) 
                                               VALUES (@T, @C, @PC, @PV, @Opt) 
                                               ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                                
                                using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                    cmd.Parameters.AddWithValue("@T", tbName); cmd.Parameters.AddWithValue("@C", colName);
                                    cmd.Parameters.AddWithValue("@PC", parentCol); cmd.Parameters.AddWithValue("@PV", parentVal);
                                    cmd.Parameters.AddWithValue("@Opt", optsStr); cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        trans.Commit();
                    }
                }
                MessageBox.Show("選項設定已儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshConfiguredCache(); 
                _cboDb.Invalidate(); _cboTable.Invalidate(); foreach(var c in _cboCols) c.Invalidate(); 
                LoadDropdownConfigs(); 
                UpdateChildParentVals(1, _dgvOptions[0]);
            } catch (Exception ex) {
                MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統下拉選單設定_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TableName AS [資料表名稱], ColName AS [欄位名稱], ParentColName AS [父層欄位], ParentValue AS [觸發條件], Options AS [選項內容(純文字)] FROM DropdownConfigs", conn))
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

                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("下拉選單設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！\n注意：Excel匯出功能不支援夾帶圖示，匯出的檔案僅包含純文字選項。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e) {
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
                                        string tb = ws.Cells[r, 1].Text.Trim(); string col = ws.Cells[r, 2].Text.Trim();
                                        string pCol = ws.Cells[r, 3].Text.Trim(); string pVal = ws.Cells[r, 4].Text.Trim(); string opt = ws.Cells[r, 5].Text.Trim();
                                        if (string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) continue;

                                        string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) VALUES (@T, @C, @PC, @PV, @Opt) ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@T", tb); cmd.Parameters.AddWithValue("@C", col);
                                            cmd.Parameters.AddWithValue("@PC", pCol); cmd.Parameters.AddWithValue("@PV", pVal); cmd.Parameters.AddWithValue("@Opt", opt);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("下拉選單設定已批次匯入並【自動存檔覆寫】成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        RefreshConfiguredCache(); LoadDropdownConfigs();
                        _cboDb.SelectedIndex = 0; _cboDb.Invalidate(); _cboTable.Invalidate(); foreach(var c in _cboCols) c.Invalidate();
                    } catch (Exception ex) { MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        public static void LoadDropdownConfigs() {
            DropdownCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM DropdownConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string tb = reader["TableName"].ToString(); string col = reader["ColName"].ToString();
                            string pCol = reader["ParentColName"].ToString(); string pVal = reader["ParentValue"].ToString(); string opts = reader["Options"].ToString();
                            string key = $"{tb}|{col}|{pCol}|{pVal}";
                            
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
                            DropdownCache[key] = defs;
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetOptions(string tableName, string colName, string parentColName = "", string parentValue = "") {
            string key = $"{tableName}|{colName}|{parentColName}|{parentValue}";
            if (DropdownCache.ContainsKey(key)) return DropdownCache[key].Select(d => d.Text).ToArray();
            return null;
        }

        public static string[] GetAllOptionsForColumn(string tableName, string colName) {
            HashSet<string> allOpts = new HashSet<string> { "" };
            foreach(var kvp in DropdownCache) {
                var parts = kvp.Key.Split('|');
                if(parts.Length == 4 && parts[0] == tableName && parts[1] == colName) {
                    foreach(var opt in kvp.Value) allOpts.Add(opt.Text);
                }
            }
            return allOpts.ToArray();
        }

        private void DgvOptions_CellContentClick(object sender, DataGridViewCellEventArgs e, int gridIndex, bool isMulti)
        {
            if (e.RowIndex < 0) return;
            DataGridView dgv = isMulti ? _dgvOptionsMulti : _dgvOptions[gridIndex];

            if (dgv.Columns[e.ColumnIndex].Name == "Upload")
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "圖片檔案 (*.png;*.jpg;*.jpeg;*.ico)|*.png;*.jpg;*.jpeg;*.ico", Title = "選擇小圖示 (建議正方形)" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (Image img = Image.FromFile(ofd.FileName))
                            {
                                using (Bitmap bmp = new Bitmap(24, 24))
                                {
                                    using (Graphics g = Graphics.FromImage(bmp))
                                    {
                                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                        g.SmoothingMode = SmoothingMode.AntiAlias;
                                        g.DrawImage(img, 0, 0, 24, 24);
                                    }
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        string base64 = Convert.ToBase64String(ms.ToArray());
                                        dgv.Rows[e.RowIndex].Cells["Icon"].Tag = base64; 
                                        dgv.Rows[e.RowIndex].Cells["Icon"].Value = Image.FromStream(new MemoryStream(ms.ToArray())); 
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("圖片處理失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            else if (dgv.Columns[e.ColumnIndex].Name == "Clear")
            {
                dgv.Rows[e.RowIndex].Cells["Icon"].Tag = null;
                dgv.Rows[e.RowIndex].Cells["Icon"].Value = null;
            }
        }
    }
}
