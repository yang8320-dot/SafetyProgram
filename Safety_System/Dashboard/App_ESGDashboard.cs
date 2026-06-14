/// FILE: Safety_System/Dashboard/App_ESGDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ESGDashboard
    {
        private ComboBox _cboYear;
        private Button _btnSearch;
        private Button _btnPdf;
        private Button _btnSettings;
        private Panel _mainScrollPanel;

        // 定義要顯示的五個資料表與對應 UI 容器
        private class SectionInfo
        {
            public string TableName { get; set; }
            public string Title { get; set; }
            public Color ThemeColor { get; set; }
            public Panel MainBox { get; set; }
            public DataGridView Dgv { get; set; }
            
            public Button BtnFilter { get; set; }
            public Label LblFilterStatus { get; set; }
        }

        private List<SectionInfo> _sections;

        private const string DbName = "ESG";
        
        // 🟢 替換：移除 TXT 路徑，改用 SQLite 儲存
        // 欄位顯示/隱藏的快取
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        // 資料列 (筆) 顯示/隱藏的快取 (紀錄被隱藏的項目名稱)
        private Dictionary<string, HashSet<string>> _hiddenRows = new Dictionary<string, HashSet<string>>();

        // 您指定的預設顯示欄位名單
        private readonly string[] _defaultVisibleCols = { 
            "年度", "部門", "國際指標", "ESG領域", "指標分類", 
            "預防投入/預期改善", "指標名稱", "實際數據呈現", "計算公式",
            // 針對 ESG_Performance 表格的預設欄位
            "年月", "單位", "項目", "說明", "預計執行週期", "費用TWD", "統計至12月底之實際數據含計算式"
        };

        // 🟢 新增：資料庫初始化邏輯
        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("CREATE TABLE IF NOT EXISTS [ESGDashboardVisibility] (ItemKey TEXT PRIMARY KEY, IsVisible INTEGER);", conn)) {
                        cmd.ExecuteNonQuery();
                    }
                    using (var cmd = new SQLiteCommand("CREATE TABLE IF NOT EXISTS [ESGDashboardHiddenRows] (TableName TEXT, ItemName TEXT, UNIQUE(TableName, ItemName));", conn)) {
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public Control GetView()
        {
            InitDatabase(); // 🟢 初始化資料表
            LoadVisibilitySettings();
            LoadHiddenRowsSettings(); // 載入資料列隱藏名單

            _sections = new List<SectionInfo>
            {
                new SectionInfo { TableName = "ESG_Performance", Title = "ESG績效管理", ThemeColor = Color.DarkOliveGreen },
                new SectionInfo { TableName = "ESG_OccupationalSafety", Title = "職業安全指標", ThemeColor = Color.SteelBlue },
                new SectionInfo { TableName = "ESG_HealthHygiene", Title = "健康衛生指標", ThemeColor = Color.Chocolate },
                new SectionInfo { TableName = "ESG_EnvironmentClimate", Title = "環境與氣侯指標", ThemeColor = Color.SeaGreen },
                new SectionInfo { TableName = "ESG_FireResilience", Title = "消防與韌性指標", ThemeColor = Color.IndianRed }
            };

            _mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            
            TableLayoutPanel masterLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            // ==========================================
            // 1. 標題與操作區
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "🌱 ESG 永續發展管理與績效看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkOliveGreen, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblTitle);

            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 10, 0, 20), Margin = new Padding(0), WrapContents = false };
            
            _cboYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 100, Margin = new Padding(0, 4, 10, 0) };
            int currYear = DateTime.Today.Year;
            for (int i = currYear - 10; i <= currYear + 2; i++) {
                _cboYear.Items.Add(i.ToString());
            }
            _cboYear.SelectedItem = currYear.ToString();

            int btnHeight = 35;

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(130, btnHeight), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            _btnSettings = new Button { Text = "⚙️ 顯示設定", Size = new Size(130, btnHeight), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnSettings.FlatAppearance.BorderSize = 0;
            _btnSettings.Click += (s, e) => OpenSettingsDialog();

            _btnPdf = new Button { Text = "📄 選擇並導出 PDF", Size = new Size(180, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnPdf.FlatAppearance.BorderSize = 0;
            _btnPdf.Click += (s, e) => ExportToPdf();

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢年度:", AutoSize = true, Margin = new Padding(0, 8, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboYear,
                _btnSearch, _btnSettings, _btnPdf
            });

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);

            // ==========================================
            // 2. 建立 5 大資料區塊
            // ==========================================
            for (int i = 0; i < _sections.Count; i++)
            {
                var sec = _sections[i];
                sec.MainBox = BuildSectionBox(sec, sec.ThemeColor);
                sec.Dgv = CreateDataGrid(sec.ThemeColor);
                
                sec.MainBox.Controls.Add(sec.Dgv);
                sec.Dgv.BringToFront(); 

                masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                masterLayout.Controls.Add(sec.MainBox, 0, i + 2);
            }

            _mainScrollPanel.Controls.Add(masterLayout);

            _ = LoadDashboardDataAsync(); // 初始載入

            return _mainScrollPanel;
        }

        private Panel BuildSectionBox(SectionInfo sec, Color themeColor)
        {
            Panel pnlBox = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 25), Padding = new Padding(15) };
            pnlBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Panel pnlTitleBar = new Panel { Dock = DockStyle.Top, Height = 45, Margin = new Padding(0, 0, 0, 10) };

            Label lblTitle = new Label { 
                Text = $"■ {sec.Title}", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                ForeColor = themeColor, 
                Dock = DockStyle.Left, 
                AutoSize = true,
                Padding = new Padding(0, 5, 0, 0)
            };

            // 資料列顯示篩選按鈕
            Button btnFilter = new Button {
                Text = "🔍 顯示資料",
                Size = new Size(130, 35),
                Dock = DockStyle.Right,
                BackColor = Color.LightSlateGray,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
            btnFilter.FlatAppearance.BorderSize = 0;
            btnFilter.Click += (s, e) => {
                if (_cboYear.SelectedItem != null) {
                    OpenRowVisibilityDialog(sec, _cboYear.SelectedItem.ToString());
                }
            };
            
            // 綁定到 sec 中，方便匯出時隱藏
            sec.BtnFilter = btnFilter;

            Label lblFilterStatus = new Label {
                Text = "",
                Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                ForeColor = Color.DarkOrange,
                Dock = DockStyle.Right,
                AutoSize = true,
                Padding = new Padding(0, 12, 10, 0)
            };
            
            // 綁定到 sec 中，方便匯出時隱藏
            sec.LblFilterStatus = lblFilterStatus;
            
            pnlTitleBar.Paint += (s, e) => {
                if (_hiddenRows.ContainsKey(sec.TableName) && _hiddenRows[sec.TableName].Count > 0) {
                    lblFilterStatus.Text = "(已隱藏部分資料列)";
                } else {
                    lblFilterStatus.Text = "";
                }
            };

            pnlTitleBar.Controls.Add(lblTitle);
            pnlTitleBar.Controls.Add(btnFilter);
            pnlTitleBar.Controls.Add(lblFilterStatus);

            pnlBox.Controls.Add(pnlTitleBar);
            return pnlBox;
        }

        private DataGridView CreateDataGrid(Color headerColor)
        {
            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Top, 
                Height = 250, 
                BackgroundColor = Color.WhiteSmoke, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                // 🟢 確保行高會因為文字太多而自動長高
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                Font = new Font("Microsoft JhengHei UI", 11F),
                BorderStyle = BorderStyle.None,
                Margin = new Padding(0, 10, 0, 0)
            };
            
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = headerColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            
            // 🟢 釋放表頭高度鎖定，允許表頭文字自動換行
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True; // 內容自動換行
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            return dgv;
        }

        // ==========================================
        // 取得資料表的代表名稱欄位
        // ==========================================
        private string GetKeyColumn(DataTable dt)
        {
            if (dt.Columns.Contains("指標名稱")) return "指標名稱";
            if (dt.Columns.Contains("項目")) return "項目";
            return ""; 
        }

        private async Task LoadDashboardDataAsync()
        {
            if (_cboYear.SelectedItem == null) return;
            string targetYear = _cboYear.SelectedItem.ToString();

            _btnSearch.Enabled = false;
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try
            {
                await Task.Run(() =>
                {
                    foreach (var sec in _sections)
                    {
                        DataTable dt = null;
                        try {
                            dt = DataManager.GetTableData(DbName, sec.TableName, "", "", "");
                        } catch { continue; }

                        if (dt != null)
                        {
                            DataView dv = dt.DefaultView;
                            
                            // 區分年度與年月欄位過濾邏輯
                            if (dt.Columns.Contains("年度")) {
                                dv.RowFilter = $"[年度] = '{targetYear}' OR [年度] = '{targetYear}年'";
                            } else if (dt.Columns.Contains("年月")) {
                                dv.RowFilter = $"[年月] LIKE '{targetYear}-%' OR [年月] LIKE '{targetYear}/%'";
                            }

                            DataTable filteredDt = dv.ToTable();

                            // 讀取該區塊被隱藏的資料列，從顯示資料表中剔除
                            string keyCol = GetKeyColumn(filteredDt);
                            if (_hiddenRows.ContainsKey(sec.TableName))
                            {
                                var hiddenSet = _hiddenRows[sec.TableName];
                                // 從後面刪除避免 Index 跑掉
                                for (int i = filteredDt.Rows.Count - 1; i >= 0; i--)
                                {
                                    DataRow r = filteredDt.Rows[i];
                                    string val = "";
                                    
                                    if (!string.IsNullOrEmpty(keyCol) && filteredDt.Columns.Contains(keyCol)) {
                                        val = r[keyCol]?.ToString().Trim() ?? "";
                                    }

                                    // 防呆：如果名稱為空，以 ID 作為識別顯示
                                    if (string.IsNullOrEmpty(val)) {
                                        string id = filteredDt.Columns.Contains("Id") ? r["Id"].ToString() : "";
                                        val = $"[未填寫名稱] (系統代碼:{id})";
                                    }

                                    if (hiddenSet.Contains(val))
                                    {
                                        filteredDt.Rows.RemoveAt(i);
                                    }
                                }
                            }

                            if (sec.Dgv.InvokeRequired) {
                                sec.Dgv.Invoke(new Action(() => BindDataToGrid(sec, filteredDt)));
                            } else {
                                BindDataToGrid(sec, filteredDt);
                            }
                        }
                    }
                });
            }
            finally
            {
                _btnSearch.Enabled = true;
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void BindDataToGrid(SectionInfo sec, DataTable dt)
        {
            sec.Dgv.DataSource = dt;

            // 初始化欄位隱藏設定
            foreach (DataGridViewColumn col in sec.Dgv.Columns)
            {
                string dictKey = $"{sec.TableName}::{col.Name}"; 

                if (_columnVisibility.ContainsKey(dictKey)) {
                    col.Visible = _columnVisibility[dictKey];
                } else {
                    // 如果沒有存過設定，套用預設顯示邏輯
                    if (_defaultVisibleCols.Contains(col.Name)) {
                        col.Visible = true;
                        _columnVisibility[dictKey] = true;
                    } else {
                        col.Visible = false;
                        _columnVisibility[dictKey] = false;
                    }
                }
            }
            
            // 動態精準計算 Grid 高度，加上小緩衝避免卷軸出現
            sec.Dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

            int totalHeight = sec.Dgv.ColumnHeadersHeight;
            foreach (DataGridViewRow row in sec.Dgv.Rows) {
                totalHeight += row.Height;
            }
            
            sec.Dgv.Height = totalHeight > 500 ? 500 : (totalHeight < 120 ? 120 : totalHeight + 5);
            sec.Dgv.ClearSelection();
            
            sec.MainBox.Invalidate(true);
        }

        // ==========================================
        // 🟢 改為 SQLite：資料列(筆)顯示/隱藏管理系統
        // ==========================================
        private void LoadHiddenRowsSettings()
        {
            _hiddenRows.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM ESGDashboardHiddenRows", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string tbl = reader["TableName"].ToString();
                            string item = reader["ItemName"].ToString();
                            if (!_hiddenRows.ContainsKey(tbl)) _hiddenRows[tbl] = new HashSet<string>();
                            _hiddenRows[tbl].Add(item);
                        }
                    }
                }
            } catch { }
        }

        private void SaveHiddenRowsSettings()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        new SQLiteCommand("DELETE FROM ESGDashboardHiddenRows", conn, trans).ExecuteNonQuery();
                        
                        string sql = "INSERT INTO ESGDashboardHiddenRows (TableName, ItemName) VALUES (@T, @I)";
                        foreach (var kvp in _hiddenRows) {
                            foreach (var item in kvp.Value) {
                                using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                    cmd.Parameters.AddWithValue("@T", kvp.Key);
                                    cmd.Parameters.AddWithValue("@I", item);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }

        private void OpenRowVisibilityDialog(SectionInfo sec, string targetYear)
        {
            // 先從 DB 撈取該年度的所有完整資料 (不套用隱藏過濾)
            DataTable dt = DataManager.GetTableData(DbName, sec.TableName, "", "", "");
            if (dt == null) return;
            
            DataView dv = dt.DefaultView;
            if (dt.Columns.Contains("年度")) {
                dv.RowFilter = $"[年度] = '{targetYear}' OR [年度] = '{targetYear}年'";
            } else if (dt.Columns.Contains("年月")) {
                dv.RowFilter = $"[年月] LIKE '{targetYear}-%' OR [年月] LIKE '{targetYear}/%'";
            }
            
            DataTable yearDt = dv.ToTable();
            string keyCol = GetKeyColumn(yearDt);
            
            HashSet<string> uniqueItems = new HashSet<string>();
            foreach (DataRow r in yearDt.Rows) {
                string val = "";
                if (!string.IsNullOrEmpty(keyCol) && yearDt.Columns.Contains(keyCol)) {
                    val = r[keyCol]?.ToString().Trim() ?? "";
                }
                
                if (string.IsNullOrEmpty(val)) {
                    string id = yearDt.Columns.Contains("Id") ? r["Id"].ToString() : "";
                    val = $"[未填寫名稱] (系統代碼:{id})";
                }
                
                uniqueItems.Add(val);
            }

            if (uniqueItems.Count == 0) {
                MessageBox.Show($"目前 {targetYear} 年度沒有任何資料可供設定。");
                return;
            }

            using (Form f = new Form { Text = $"🔍 {sec.Title} - 顯示資料選擇", Size = new Size(500, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.WhiteSmoke })
            {
                // 🟢 限制標籤的最大寬度，強迫它在視窗邊緣自動換行
                Label lblTop = new Label { 
                    Text = "請勾選您希望顯示在看板上的資料，取消勾選則會隱藏：", 
                    Dock = DockStyle.Top, 
                    Padding = new Padding(15), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = sec.ThemeColor, 
                    AutoSize = true,
                    MaximumSize = new Size(470, 0)
                };
                f.Controls.Add(lblTop);

                CheckedListBox clb = new CheckedListBox { 
                    Dock = DockStyle.Fill, 
                    Font = new Font("Microsoft JhengHei UI", 12F), 
                    CheckOnClick = true, 
                    BorderStyle = BorderStyle.None,
                    Margin = new Padding(20)
                };

                bool hasHiddenSet = _hiddenRows.ContainsKey(sec.TableName);
                foreach (string item in uniqueItems) {
                    bool isVisible = true;
                    if (hasHiddenSet && _hiddenRows[sec.TableName].Contains(item)) {
                        isVisible = false;
                    }
                    clb.Items.Add(item, isVisible);
                }
                
                f.Controls.Add(clb);
                clb.BringToFront();

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(15, 0, 15, 10) };
                
                Button btnSave = new Button { 
                    Text = "💾 儲存並套用顯示設定", 
                    Dock = DockStyle.Fill, 
                    BackColor = Color.ForestGreen, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat 
                };
                btnSave.FlatAppearance.BorderSize = 0;
                
                btnSave.Click += (s, e) => {
                    if (!_hiddenRows.ContainsKey(sec.TableName)) {
                        _hiddenRows[sec.TableName] = new HashSet<string>();
                    } else {
                        _hiddenRows[sec.TableName].Clear(); 
                    }

                    for (int i = 0; i < clb.Items.Count; i++) {
                        if (!clb.GetItemChecked(i)) {
                            _hiddenRows[sec.TableName].Add(clb.Items[i].ToString());
                        }
                    }

                    SaveHiddenRowsSettings();
                    _ = LoadDashboardDataAsync();
                    f.DialogResult = DialogResult.OK;
                };

                pnlBottom.Controls.Add(btnSave);
                f.Controls.Add(pnlBottom);

                f.ShowDialog();
            }
        }

        // ==========================================
        // 🟢 改為 SQLite：欄位 (直行) 顯示設定系統
        // ==========================================
        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM ESGDashboardVisibility", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            _columnVisibility[reader["ItemKey"].ToString()] = Convert.ToInt32(reader["IsVisible"]) == 1;
                        }
                    }
                }
            } catch { }
        }

        private void SaveVisibilitySettings()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        new SQLiteCommand("DELETE FROM ESGDashboardVisibility", conn, trans).ExecuteNonQuery();
                        
                        string sql = "INSERT INTO ESGDashboardVisibility (ItemKey, IsVisible) VALUES (@K, @V)";
                        foreach (var kvp in _columnVisibility) {
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@K", kvp.Key);
                                cmd.Parameters.AddWithValue("@V", kvp.Value ? 1 : 0);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }

        private void OpenSettingsDialog()
        {
            using (Form f = new Form { Text = "⚙️ 顯示設定", Size = new Size(800, 550), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.WhiteSmoke }) 
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));

                Label lblTop = new Label { 
                    Text = "請設定查詢看板時【允許顯示】的資料表欄位：", 
                    Dock = DockStyle.Fill, 
                    Padding = new Padding(15, 15, 10, 5), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = Color.SteelBlue, 
                    AutoSize = true,
                    MaximumSize = new Size(770, 0)
                };
                tlp.Controls.Add(lblTop, 0, 0);

                // 🟢 改用保證等比縮放的 TableLayoutPanel (左40% : 右60%)
                TableLayoutPanel splitLayout = new TableLayoutPanel {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 1,
                    Margin = new Padding(15, 5, 15, 15)
                };
                splitLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
                splitLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));

                // 左半部
                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 0, 5, 0) };
                Label lblLeft = new Label { Text = "1. 選擇看板區塊", Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Top, Height = 25 };
                ListBox lbTables = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), IntegralHeight = false };
                foreach (var sec in _sections) lbTables.Items.Add(sec.Title);
                pnlLeft.Controls.Add(lbTables);
                pnlLeft.Controls.Add(lblLeft);

                // 右半部
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5, 0, 0, 0) };
                Label lblRight = new Label { Text = "2. 勾選要顯示的欄位", Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Top, Height = 25 };
                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false };
                pnlRight.Controls.Add(clbCols);
                pnlRight.Controls.Add(lblRight);

                splitLayout.Controls.Add(pnlLeft, 0, 0);
                splitLayout.Controls.Add(pnlRight, 1, 0);
                
                tlp.Controls.Add(splitLayout, 0, 1);

                Button btnSave = new Button { 
                    Text = "💾 儲存並套用", 
                    Dock = DockStyle.Fill, 
                    BackColor = Color.ForestGreen, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat,
                    Margin = new Padding(0)
                };
                btnSave.FlatAppearance.BorderSize = 0;
                
                Panel pnlBottom = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15, 0, 15, 10) };
                pnlBottom.Controls.Add(btnSave);
                tlp.Controls.Add(pnlBottom, 0, 2);

                f.Controls.Add(tlp);

                lbTables.SelectedIndexChanged += (s, e) => {
                    if (lbTables.SelectedIndex < 0) return;
                    clbCols.Items.Clear();
                    string tblName = _sections[lbTables.SelectedIndex].TableName;
                    var cols = DataManager.GetColumnNames(DbName, tblName);

                    foreach (var c in cols) {
                        if (c == "Id") continue;
                        string key = $"{tblName}::{c}";
                        bool isChecked = _columnVisibility.ContainsKey(key) ? _columnVisibility[key] : _defaultVisibleCols.Contains(c);
                        clbCols.Items.Add(c, isChecked);
                    }
                };

                clbCols.ItemCheck += (s, e) => {
                    if (lbTables.SelectedIndex < 0) return;
                    string tblName = _sections[lbTables.SelectedIndex].TableName;
                    string colName = clbCols.Items[e.Index].ToString();
                    _columnVisibility[$"{tblName}::{colName}"] = e.NewValue == CheckState.Checked;
                };

                btnSave.Click += (s, e) => {
                    SaveVisibilitySettings();
                    _ = LoadDashboardDataAsync();
                    f.DialogResult = DialogResult.OK;
                };

                if (lbTables.Items.Count > 0) lbTables.SelectedIndex = 0;

                f.ShowDialog();
            }
        }

        // ==========================================
        // PDF 導出系統
        // ==========================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 450, Height = 400, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));

                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表區塊：", Dock = DockStyle.Fill, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), AutoSize = true };
                tlp.Controls.Add(lbl, 0, 0);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(15, 5, 15, 5), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
                
                foreach (var sec in _sections) {
                    clb.Items.Add(sec.Title, true); 
                }
                tlp.Controls.Add(clb, 0, 1);

                Panel pnlBottom = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) };
                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                pnlBottom.Controls.Add(btnOk);
                tlp.Controls.Add(pnlBottom, 0, 2);

                f.Controls.Add(tlp);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    for (int i = 0; i < clb.Items.Count; i++) {
                        if (clb.GetItemChecked(i)) {
                            selectedPanels.Add(_sections[i].MainBox);
                        }
                    }
                }
            }
            return selectedPanels;
        }

        private void ExportToPdf()
        {
            var panelsToExport = GetSelectedExportPanels();
            if (panelsToExport.Count == 0) return;

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try 
            {
                Application.DoEvents(); 

                List<Bitmap> bitmaps = new List<Bitmap>();
                foreach (var pnl in panelsToExport) 
                {
                    DataGridView dgv = pnl.Controls.OfType<DataGridView>().FirstOrDefault();
                    
                    SectionInfo sec = _sections.FirstOrDefault(s => s.MainBox == pnl);
                    bool origBtnVisible = true;
                    bool origLblVisible = true;
                    
                    if (sec != null) {
                        origBtnVisible = sec.BtnFilter.Visible;
                        origLblVisible = sec.LblFilterStatus.Visible;
                        sec.BtnFilter.Visible = false;
                        sec.LblFilterStatus.Visible = false;
                    }

                    int originalHeight = pnl.Height;
                    int dgvOriginalHeight = 0;

                    if (dgv != null) {
                        dgvOriginalHeight = dgv.Height;
                        int totalHeight = dgv.ColumnHeadersHeight;
                        foreach (DataGridViewRow row in dgv.Rows) totalHeight += row.Height;
                        dgv.Height = totalHeight + 2; 
                        pnl.Height = dgv.Height + 50; 
                    }

                    Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                    pnl.DrawToBitmap(bmp, new Rectangle(0, 0, pnl.Width, pnl.Height));
                    bitmaps.Add(bmp);

                    // 恢復狀態
                    if (dgv != null) {
                        dgv.Height = dgvOriginalHeight;
                        pnl.Height = originalHeight;
                    }
                    if (sec != null) {
                        sec.BtnFilter.Visible = origBtnVisible;
                        sec.LblFilterStatus.Visible = origLblVisible;
                    }
                }

                string dateStr = $"查詢年度：{_cboYear.SelectedItem}";
                PdfHelper.ExportDashboardToPdf(bitmaps, "ESG 永續發展管理與績效報表", dateStr, "ESG永續績效報表");
            } 
            catch (Exception ex)
            {
                MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }
    }
}
