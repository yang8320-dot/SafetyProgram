/// FILE: Safety_System/Dashboard/App_ESGDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
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
        }

        private List<SectionInfo> _sections;

        private const string DbName = "ESG";
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ESGDashboard_Visibility.txt");
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        // 🟢 新增：各資料表的進階資料篩選記憶設定
        private readonly string FilterFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ESGDashboard_Filter.txt");
        private Dictionary<string, Dictionary<string, string>> _tableFilters = new Dictionary<string, Dictionary<string, string>>();

        // 您指定的預設顯示欄位名單
        private readonly string[] _defaultVisibleCols = { 
            "年度", "部門", "國際指標", "ESG領域", "指標分類", 
            "預防投入/預期改善", "指標名稱", "實際數據呈現", "計算公式",
            // 針對 ESG_Performance 表格的預設欄位
            "年月", "單位", "項目", "說明", "預計執行週期", "費用TWD", "統計至12月底之實際數據含計算式"
        };

        public Control GetView()
        {
            LoadVisibilitySettings();
            LoadFilterSettings(); // 🟢 載入篩選記憶

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

        // 🟢 修改點：傳入整個 SectionInfo 物件，並在右上角加入「顯示資料(篩選)」按鈕
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

            Button btnFilter = new Button {
                Text = "🔍 顯示資料 (篩選)",
                Size = new Size(160, 35),
                Dock = DockStyle.Right,
                BackColor = Color.LightSlateGray,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
            btnFilter.FlatAppearance.BorderSize = 0;
            btnFilter.Click += (s, e) => OpenFilterDialog(sec);

            // 判斷該區塊目前是否有啟用過濾條件，有的話加上提示
            Label lblFilterStatus = new Label {
                Text = "",
                Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                ForeColor = Color.DarkOrange,
                Dock = DockStyle.Right,
                AutoSize = true,
                Padding = new Padding(0, 12, 10, 0)
            };
            
            // 每次畫面刷新前更新這個狀態文字
            pnlTitleBar.Paint += (s, e) => {
                if (_tableFilters.ContainsKey(sec.TableName) && _tableFilters[sec.TableName].Any(kvp => !string.IsNullOrWhiteSpace(kvp.Value))) {
                    lblFilterStatus.Text = "(已啟用進階篩選)";
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
            dgv.ColumnHeadersHeight = 40;
            
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            return dgv;
        }

        // ==========================================
        // 🟢 載入與套用篩選資料
        // ==========================================
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
                            List<string> filterConditions = new List<string>();
                            
                            // 1. 基礎年度過濾
                            if (dt.Columns.Contains("年度")) {
                                filterConditions.Add($"([年度] = '{targetYear}' OR [年度] = '{targetYear}年')");
                            } else if (dt.Columns.Contains("年月")) {
                                filterConditions.Add($"([年月] LIKE '{targetYear}-%' OR [年月] LIKE '{targetYear}/%')");
                            }

                            // 2. 🟢 套用自訂的顯示過濾條件 (來自記憶設定)
                            if (_tableFilters.ContainsKey(sec.TableName))
                            {
                                foreach (var kvp in _tableFilters[sec.TableName])
                                {
                                    if (dt.Columns.Contains(kvp.Key) && !string.IsNullOrWhiteSpace(kvp.Value))
                                    {
                                        // 使用 LIKE 進行模糊比對，並跳脫單引號以防報錯
                                        filterConditions.Add($"[{kvp.Key}] LIKE '%{kvp.Value.Replace("'", "''")}%'");
                                    }
                                }
                            }

                            dv.RowFilter = filterConditions.Count > 0 ? string.Join(" AND ", filterConditions) : "";
                            DataTable filteredDt = dv.ToTable();

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
            
            // 動態調整 Grid 高度以適應內容，避免出現卷軸 (供匯出 PDF 使用)
            int totalHeight = sec.Dgv.ColumnHeadersHeight;
            foreach (DataGridViewRow row in sec.Dgv.Rows) {
                totalHeight += row.Height;
            }
            
            sec.Dgv.Height = totalHeight > 500 ? 500 : (totalHeight < 150 ? 150 : totalHeight + 2);
            sec.Dgv.ClearSelection();
            
            // 觸發重新繪製以更新標題旁的「(已啟用進階篩選)」文字
            sec.MainBox.Invalidate(true);
        }

        // ==========================================
        // 🟢 專屬各表框的「顯示資料 (進階篩選)」管理系統
        // ==========================================
        private void LoadFilterSettings()
        {
            _tableFilters.Clear();
            if (File.Exists(FilterFile)) {
                try {
                    foreach (var line in File.ReadAllLines(FilterFile, Encoding.UTF8)) {
                        var parts = line.Split(new[] { "::" }, StringSplitOptions.None);
                        if (parts.Length == 3) {
                            string tbl = parts[0];
                            string col = parts[1];
                            string val = parts[2];
                            if (!_tableFilters.ContainsKey(tbl)) _tableFilters[tbl] = new Dictionary<string, string>();
                            _tableFilters[tbl][col] = val;
                        }
                    }
                } catch { }
            }
        }

        private void SaveFilterSettings()
        {
            try {
                List<string> lines = new List<string>();
                foreach (var tbl in _tableFilters) {
                    foreach (var col in tbl.Value) {
                        if (!string.IsNullOrWhiteSpace(col.Value)) {
                            lines.Add($"{tbl.Key}::{col.Key}::{col.Value}");
                        }
                    }
                }
                File.WriteAllLines(FilterFile, lines, Encoding.UTF8);
            } catch { }
        }

        private void OpenFilterDialog(SectionInfo sec)
        {
            using (Form f = new Form { Text = $"🔍 {sec.Title} - 資料進階篩選", Size = new Size(500, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.WhiteSmoke })
            {
                Label lblTop = new Label { 
                    Text = "請輸入各欄位的查詢關鍵字 (留空代表不限制)：", 
                    Dock = DockStyle.Top, 
                    Padding = new Padding(15), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = sec.ThemeColor, 
                    AutoSize = true 
                };
                f.Controls.Add(lblTop);

                FlowLayoutPanel flp = new FlowLayoutPanel { 
                    Dock = DockStyle.Fill, 
                    AutoScroll = true, 
                    FlowDirection = FlowDirection.TopDown, 
                    WrapContents = false, 
                    Padding = new Padding(20, 10, 20, 10) 
                };
                f.Controls.Add(flp);
                flp.BringToFront();

                List<string> cols = DataManager.GetColumnNames(DbName, sec.TableName);
                Dictionary<string, TextBox> inputMap = new Dictionary<string, TextBox>();

                // 排除系統專用或不適合做文字篩選的欄位
                string[] excludeCols = { "Id", "年度", "年月", "最後修改人", "修改時間" };

                foreach (string col in cols)
                {
                    if (excludeCols.Contains(col)) continue;

                    Panel pRow = new Panel { Width = 420, Height = 40, Margin = new Padding(0, 0, 0, 10) };
                    Label lCol = new Label { Text = col + "：", Width = 150, Location = new Point(0, 7), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleRight };
                    TextBox tInput = new TextBox { Width = 260, Location = new Point(155, 5), Font = new Font("Microsoft JhengHei UI", 12F) };

                    // 帶入之前存檔的記憶值
                    if (_tableFilters.ContainsKey(sec.TableName) && _tableFilters[sec.TableName].ContainsKey(col)) {
                        tInput.Text = _tableFilters[sec.TableName][col];
                    }

                    pRow.Controls.Add(lCol);
                    pRow.Controls.Add(tInput);
                    flp.Controls.Add(pRow);

                    inputMap[col] = tInput;
                }

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(15, 0, 15, 10) };
                
                Button btnClear = new Button { 
                    Text = "🧹 清除所有條件", 
                    Width = 150, 
                    Dock = DockStyle.Left, 
                    BackColor = Color.IndianRed, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat 
                };
                btnClear.FlatAppearance.BorderSize = 0;
                btnClear.Click += (s, e) => {
                    foreach (var t in inputMap.Values) t.Clear();
                };

                Button btnSave = new Button { 
                    Text = "💾 儲存並套用篩選", 
                    Width = 200, 
                    Dock = DockStyle.Right, 
                    BackColor = Color.ForestGreen, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat 
                };
                btnSave.FlatAppearance.BorderSize = 0;
                
                btnSave.Click += (s, e) => {
                    if (!_tableFilters.ContainsKey(sec.TableName)) {
                        _tableFilters[sec.TableName] = new Dictionary<string, string>();
                    } else {
                        _tableFilters[sec.TableName].Clear(); // 清空舊的
                    }

                    foreach (var kvp in inputMap) {
                        string val = kvp.Value.Text.Trim();
                        if (!string.IsNullOrEmpty(val)) {
                            _tableFilters[sec.TableName][kvp.Key] = val;
                        }
                    }

                    SaveFilterSettings();
                    _ = LoadDashboardDataAsync(); // 觸發重新讀取並套用篩選
                    f.DialogResult = DialogResult.OK;
                };

                pnlBottom.Controls.Add(btnClear);
                pnlBottom.Controls.Add(btnSave);
                f.Controls.Add(pnlBottom);

                f.ShowDialog();
            }
        }

        // ==========================================
        // 欄位顯示設定系統
        // ==========================================
        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            if (File.Exists(VisibilityFile)) {
                try {
                    foreach (var line in File.ReadAllLines(VisibilityFile, Encoding.UTF8)) {
                        var parts = line.Split('|');
                        if (parts.Length == 2) {
                            _columnVisibility[parts[0]] = (parts[1] == "1");
                        }
                    }
                } catch { }
            }
        }

        private void SaveVisibilitySettings()
        {
            try {
                var lines = _columnVisibility.Select(kvp => $"{kvp.Key}|{(kvp.Value ? "1" : "0")}").ToArray();
                File.WriteAllLines(VisibilityFile, lines, Encoding.UTF8);
            } catch { }
        }

        private void OpenSettingsDialog()
        {
            using (Form f = new Form { Text = "⚙️ 顯示設定", Size = new Size(650, 550), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.WhiteSmoke }) 
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));

                Label lblTop = new Label { 
                    Text = "請選擇分類並勾選查詢時【允許顯示】的欄位：", 
                    Dock = DockStyle.Fill, 
                    Padding = new Padding(10), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = Color.SteelBlue, 
                    AutoSize = true 
                };
                tlp.Controls.Add(lblTop, 0, 0);

                SplitContainer split = new SplitContainer { 
                    Dock = DockStyle.Fill, 
                    SplitterDistance = 250, 
                    FixedPanel = FixedPanel.Panel1, 
                    Padding = new Padding(10) 
                };

                ListBox lbTables = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                foreach (var sec in _sections) lbTables.Items.Add(sec.Title);
                split.Panel1.Controls.Add(lbTables);

                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle };
                split.Panel2.Controls.Add(clbCols);
                tlp.Controls.Add(split, 0, 1);

                Button btnSave = new Button { 
                    Text = "💾 儲存並套用", 
                    Dock = DockStyle.Fill, 
                    BackColor = Color.ForestGreen, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat 
                };
                tlp.Controls.Add(btnSave, 0, 2);

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
                    // 暫時將 DataGridView 高度展開以擷取完整截圖
                    DataGridView dgv = pnl.Controls.OfType<DataGridView>().FirstOrDefault();
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
                    pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, pnl.Height));
                    bitmaps.Add(bmp);

                    // 恢復原本高度
                    if (dgv != null) {
                        dgv.Height = dgvOriginalHeight;
                        pnl.Height = originalHeight;
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
