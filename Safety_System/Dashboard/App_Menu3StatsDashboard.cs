/// FILE: Safety_System/Dashboard/App_Menu3StatsDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_Menu3StatsDashboard
    {
        private string _menuDbName; 
        private Panel _mainScrollPanel;
        private TableLayoutPanel _tlpThemesContainer; 
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 資料庫表名
        private const string TblThemes = "Menu3StatsDashboard_Themes";
        private const string TblConfigs = "Menu3StatsDashboard_Configs";
        private const string TblRecords = "Menu3StatsDashboard_Records";

        // 動態區塊的 UI 封裝
        private class ThemeSectionUI
        {
            public int ThemeId { get; set; }
            public string ThemeName { get; set; }
            public Panel MainBox { get; set; } 
            public ComboBox CboStartYear, CboStartMonth, CboEndYear, CboEndMonth;
            public DataGridView Dgv { get; set; }
            public FlowLayoutPanel FlpControls { get; set; } 
        }

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName;
        }

        private class EditingConfig {
            public string Name { get; set; }
            public string Formula { get; set; }
        }

        public App_Menu3StatsDashboard(string menuDbName)
        {
            _menuDbName = menuDbName;
            _dbMap = App_DbConfig.GetDbMapCache();
        }

        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblThemes}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, MenuName TEXT, ThemeName TEXT);", conn).ExecuteNonQuery();
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblConfigs}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, ThemeId INTEGER, ItemName TEXT, FormulaTemplate TEXT);", conn).ExecuteNonQuery();
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblRecords}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, ThemeId INTEGER, PeriodStart TEXT, PeriodEnd TEXT, ItemName TEXT, DataValue TEXT, Attachment TEXT, Remarks TEXT);", conn).ExecuteNonQuery();
                }
            } catch { }
        }

        public Control GetView()
        {
            InitDatabase();

            _mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };

            TableLayoutPanel masterLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2, Margin = new Padding(0) };
            masterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ================= Header =================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 70, Margin = new Padding(0,0,0,10) };
            Label lblTitle = new Label { Text = "📊 選單 3 - 動態統計看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true };
            
            Button btnAddTheme = new Button { Text = "➕ 新增主題區塊", Size = new Size(180, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Dock = DockStyle.Right };
            btnAddTheme.FlatAppearance.BorderSize = 0;
            btnAddTheme.Click += BtnAddTheme_Click;

            Button btnGlobalPdf = new Button { Text = "📄 導出選定項目 (PDF)", Size = new Size(220, 45), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Dock = DockStyle.Right };
            btnGlobalPdf.FlatAppearance.BorderSize = 0;
            btnGlobalPdf.Click += BtnGlobalPdf_Click;

            Button btnGlobalExcel = new Button { Text = "📤 導出選定項目 (Excel)", Size = new Size(230, 45), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Dock = DockStyle.Right };
            btnGlobalExcel.FlatAppearance.BorderSize = 0;
            btnGlobalExcel.Click += BtnGlobalExcel_Click;

            pnlHeader.Controls.Add(btnAddTheme);
            pnlHeader.Controls.Add(new Panel { Width = 15, Dock = DockStyle.Right }); 
            pnlHeader.Controls.Add(btnGlobalPdf);
            pnlHeader.Controls.Add(new Panel { Width = 15, Dock = DockStyle.Right }); 
            pnlHeader.Controls.Add(btnGlobalExcel);

            pnlHeader.Controls.Add(lblTitle);
            masterLayout.Controls.Add(pnlHeader, 0, 0);

            // ================= Container =================
            _tlpThemesContainer = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1,
                Padding = new Padding(0)
            };
            _tlpThemesContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); 

            masterLayout.Controls.Add(_tlpThemesContainer, 0, 1);
            _mainScrollPanel.Controls.Add(masterLayout);

            LoadThemes();

            foreach (var ui in _sections) {
                _ = CalculateAndLoadGrid(ui);
            }

            return _mainScrollPanel;
        }

        private void BtnAddTheme_Click(object sender, EventArgs e)
        {
            string themeName = ShowInputBox("請輸入新主題統計的名稱：", "新增主題", "新主題區塊");
            if (string.IsNullOrWhiteSpace(themeName)) return;

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"INSERT INTO {TblThemes} (MenuName, ThemeName) VALUES (@M, @T)", conn)) {
                        cmd.Parameters.AddWithValue("@M", _menuDbName);
                        cmd.Parameters.AddWithValue("@T", themeName.Trim());
                        cmd.ExecuteNonQuery();
                    }
                }
                
                LoadThemes();
                
                foreach (var ui in _sections) {
                    _ = CalculateAndLoadGrid(ui);
                }
                
            } catch (Exception ex) { MessageBox.Show("新增失敗：" + ex.Message); }
        }

        private List<ThemeSectionUI> _sections = new List<ThemeSectionUI>();

        private void LoadThemes()
        {
            _tlpThemesContainer.Controls.Clear();
            _tlpThemesContainer.RowStyles.Clear();
            _sections.Clear();

            DataTable dtThemes = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"SELECT * FROM {TblThemes} WHERE MenuName=@M", conn)) {
                        cmd.Parameters.AddWithValue("@M", _menuDbName);
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtThemes);
                    }
                }
            } catch { return; }

            if (dtThemes.Rows.Count == 0) {
                _tlpThemesContainer.Controls.Add(new Label { Text = "目前沒有任何統計主題，請點擊上方按鈕新增。", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Margin = new Padding(20) });
                return;
            }

            foreach (DataRow row in dtThemes.Rows) {
                BuildThemeSection(Convert.ToInt32(row["Id"]), row["ThemeName"].ToString());
            }
        }

        private void BuildThemeSection(int themeId, string themeName)
        {
            ThemeSectionUI ui = new ThemeSectionUI { ThemeId = themeId, ThemeName = themeName };

            ui.MainBox = new Panel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 0, 20) 
            };

            // ====== 查詢與操作區 ======
            GroupBox boxQuery = new GroupBox {
                Text = $"📌 {themeName} - 查詢及操作區",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                ForeColor = Color.SaddleBrown,
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(10)
            };

            ui.FlpControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Padding = new Padding(0, 5, 0, 5), Margin = new Padding(0) };
            
            ui.CboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 90, Margin = new Padding(10, 6, 5, 0) };
            ui.CboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60, Margin = new Padding(0, 6, 5, 0) };
            ui.CboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 90, Margin = new Padding(0, 6, 5, 0) };
            ui.CboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60, Margin = new Padding(0, 6, 20, 0) };

            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) { ui.CboStartYear.Items.Add(i.ToString()); ui.CboEndYear.Items.Add(i.ToString()); }
            for (int i = 1; i <= 12; i++) { ui.CboStartMonth.Items.Add(i.ToString("D2")); ui.CboEndMonth.Items.Add(i.ToString("D2")); }
            
            DateTime prevMonth = DateTime.Today.AddMonths(-1);
            ui.CboStartYear.SelectedItem = prevMonth.Year.ToString(); 
            ui.CboStartMonth.SelectedItem = prevMonth.Month.ToString("D2");
            ui.CboEndYear.SelectedItem = prevMonth.Year.ToString(); 
            ui.CboEndMonth.SelectedItem = prevMonth.Month.ToString("D2");

            Button btnSearch = new Button { Text = "🔍 讀取", Size = new Size(100, 36), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnSearch.FlatAppearance.BorderSize = 0;
            Button btnRecalc = new Button { Text = "🔄 重新統計", Size = new Size(145, 36), BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnRecalc.FlatAppearance.BorderSize = 0;
            Button btnSave = new Button { Text = "💾 儲存", Size = new Size(100, 36), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnSave.FlatAppearance.BorderSize = 0;
            Button btnSettings = new Button { Text = "⚙️ 顯示設定", Size = new Size(130, 36), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnSettings.FlatAppearance.BorderSize = 0;
            
            Button btnDelTheme = new Button { Text = "🗑️", Size = new Size(80, 36), BackColor = Color.LightCoral, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(15, 2, 0, 0) }; btnDelTheme.FlatAppearance.BorderSize = 0;

            ui.FlpControls.Controls.AddRange(new Control[] {
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 10, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartMonth, new Label { Text = "月 ~ ", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 20, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearch, btnRecalc, btnSave, btnSettings, btnDelTheme
            });
            
            boxQuery.Controls.Add(ui.FlpControls);

            // ====== 資料表格區 ======
            GroupBox boxGrid = new GroupBox {
                Text = $"📊 {themeName} - 統計數據資料表",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                ForeColor = Color.Teal,
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(10, 15, 10, 20) // 底部保留 20px
            };

            ui.Dgv = new DataGridView {
                Dock = DockStyle.Top, Height = 100, BackgroundColor = Color.White,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToResizeColumns = true,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 12F),
                BorderStyle = BorderStyle.None, CellBorderStyle = DataGridViewCellBorderStyle.Single,
                ScrollBars = ScrollBars.None, 
                Margin = new Padding(0)
            };
            ui.Dgv.EnableHeadersVisualStyles = false;
            ui.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SlateGray;
            ui.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            ui.Dgv.ColumnHeadersHeight = 40;
            ui.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            ui.Dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            ui.Dgv.Columns.Add("項目", "項目"); ui.Dgv.Columns["項目"].ReadOnly = true; 
            ui.Dgv.Columns.Add("數據", "數據資料"); ui.Dgv.Columns["數據"].ReadOnly = true; 
            ui.Dgv.Columns.Add("附件檔案", "附件檔案"); ui.Dgv.Columns["附件檔案"].ReadOnly = true; 
            ui.Dgv.Columns["附件檔案"].DefaultCellStyle.ForeColor = Color.Blue; ui.Dgv.Columns["附件檔案"].DefaultCellStyle.Font = new Font(ui.Dgv.Font, FontStyle.Underline);
            ui.Dgv.Columns.Add("備註", "備註"); 

            ui.Dgv.ColumnWidthChanged += (s, e) => {
                if (e.Column != null && e.Column.Width > 0) {
                    DataManager.SaveGridConfig("Menu3Stats", ui.ThemeName, "Width", e.Column.Name, e.Column.Width.ToString());
                    AdjustGridHeight(ui);
                }
            };

            ui.Dgv.CellFormatting += Dgv_CellFormatting;
            ui.Dgv.CellClick += (s, e) => Dgv_CellClick(s, e, ui);

            boxGrid.Controls.Add(ui.Dgv);

            ui.MainBox.Controls.Add(boxGrid);
            ui.MainBox.Controls.Add(boxQuery);
            
            _tlpThemesContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _tlpThemesContainer.Controls.Add(ui.MainBox);

            _sections.Add(ui);

            // ====== 綁定事件 ======
            btnSearch.Click += async (s, e) => await CalculateAndLoadGrid(ui);
            btnRecalc.Click += async (s, e) => await RecalculateGridData(ui); 
            btnSave.Click += (s, e) => SaveGridData(ui);
            btnSettings.Click += (s, e) => { OpenSettingsDialog(ui); _ = CalculateAndLoadGrid(ui); };
            btnDelTheme.Click += (s, e) => {
                if (MessageBox.Show($"確定要刪除主題【{themeName}】及內部所有設定與資料嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            new SQLiteCommand($"DELETE FROM {TblThemes} WHERE Id={themeId}", conn).ExecuteNonQuery();
                            new SQLiteCommand($"DELETE FROM {TblConfigs} WHERE ThemeId={themeId}", conn).ExecuteNonQuery();
                            new SQLiteCommand($"DELETE FROM {TblRecords} WHERE ThemeId={themeId}", conn).ExecuteNonQuery();
                        }
                        LoadThemes();
                    } catch { }
                }
            };
        }

        private void AdjustGridHeight(ThemeSectionUI ui)
        {
            if (ui.Dgv == null || ui.MainBox == null) return;
            
            ui.Dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

            int dgvHeight = ui.Dgv.ColumnHeadersHeight;
            foreach (DataGridViewRow row in ui.Dgv.Rows) {
                dgvHeight += row.Height;
            }
            
            ui.Dgv.Height = dgvHeight > 0 ? dgvHeight + 2 : 50;
        }

        private void ApplyColumnWidths(ThemeSectionUI ui)
        {
            var widthDict = DataManager.LoadGridConfig("Menu3Stats", ui.ThemeName, "Width");
            
            if (ui.Dgv.Columns.Contains("項目")) ui.Dgv.Columns["項目"].Width = 200;
            if (ui.Dgv.Columns.Contains("數據")) ui.Dgv.Columns["數據"].Width = 350;
            if (ui.Dgv.Columns.Contains("附件檔案")) ui.Dgv.Columns["附件檔案"].Width = 200;
            if (ui.Dgv.Columns.Contains("備註")) ui.Dgv.Columns["備註"].Width = 300;

            foreach (var kvp in widthDict) {
                if (ui.Dgv.Columns.Contains(kvp.Key)) {
                    if (int.TryParse(kvp.Value, out int w) && w > 0) {
                        ui.Dgv.Columns[kvp.Key].Width = w;
                    }
                }
            }
        }

        private string EvaluateStatsFormula(string template, string startYM, string endYM, Dictionary<string, DataTable> tableCache)
        {
            string evalText = template;
            
            Regex crossRegex = new Regex(@"(?<agg>SUM|AVG|MAX|MIN|COUNT)\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\](?:\.\[(?<dateCol>[^\]]+)\])?(?:\.\[(?<refCol>[^\]]+)\]\.\[(?<filterVal>[^\]]*)\])?\)");
            Regex mathBlockRegex = new Regex(@"\{(?<expr>[^\}]+)\}");

            var mathBlocks = mathBlockRegex.Matches(evalText);
            foreach (Match mBlock in mathBlocks)
            {
                string expr = mBlock.Groups["expr"].Value;
                var crossMatches = crossRegex.Matches(expr);
                
                foreach (Match mCross in crossMatches) {
                    string agg = mCross.Groups["agg"].Value;
                    string fDb = mCross.Groups["db"].Value;
                    string fTb = mCross.Groups["tb"].Value;
                    string fCol = mCross.Groups["col"].Value;
                    
                    string formulaDateCol = mCross.Groups["dateCol"].Success ? mCross.Groups["dateCol"].Value : ""; 
                    if (formulaDateCol == "自動判斷") formulaDateCol = "";

                    string refCol = mCross.Groups["refCol"].Success ? mCross.Groups["refCol"].Value : "";
                    string filterVal = mCross.Groups["filterVal"].Success ? mCross.Groups["filterVal"].Value : "";

                    string cacheKey = $"{fDb}|{fTb}";
                    if (!tableCache.ContainsKey(cacheKey)) {
                        tableCache[cacheKey] = DataManager.GetTableData(fDb, fTb, "", "", "");
                    }

                    double computedVal = 0;
                    DataTable fDt = tableCache[cacheKey];
                    
                    if (fDt != null) {
                        string dateCol = "";
                        if (!string.IsNullOrEmpty(formulaDateCol) && fDt.Columns.Contains(formulaDateCol)) {
                            dateCol = formulaDateCol;
                        } else {
                            dateCol = fDt.Columns.Contains("日期") ? "日期" :
                                      (fDt.Columns.Contains("清運日期") ? "清運日期" :
                                      (fDt.Columns.Contains("年月") ? "年月" :
                                      (fDt.Columns.Contains("年度") ? "年度" : "")));
                        }

                        var matchedRows = fDt.Rows.Cast<DataRow>().Where(r => {
                            if (r.RowState == DataRowState.Deleted) return false;
                            
                            if (!string.IsNullOrEmpty(dateCol)) {
                                string dVal = r[dateCol]?.ToString().Trim() ?? "";
                                if (string.IsNullOrEmpty(dVal)) return false;
                                
                                if (dateCol == "年度") {
                                    string yearVal = dVal.Replace("年", "").Trim();
                                    string sYear = startYM.Substring(0, 4);
                                    string eYear = endYM.Substring(0, 4);
                                    if (string.Compare(yearVal, sYear) < 0 || string.Compare(yearVal, eYear) > 0) return false;
                                }
                                else {
                                    dVal = dVal.Replace("/", "-");
                                    string compYm = "";
                                    var parts = dVal.Split('-');
                                    if (parts.Length >= 2) {
                                        compYm = $"{parts[0]}-{parts[1].PadLeft(2, '0')}";
                                    } else if (dVal.Length == 6 && !dVal.Contains("-")) { 
                                        compYm = $"{dVal.Substring(0,4)}-{dVal.Substring(4,2)}";
                                    } else {
                                        compYm = dVal.Length >= 7 ? dVal.Substring(0, 7) : dVal;
                                    }
                                    if (string.Compare(compYm, startYM) < 0 || string.Compare(compYm, endYM) > 0) return false;
                                }
                            }

                            if (!string.IsNullOrEmpty(refCol) && fDt.Columns.Contains(refCol)) {
                                string rowRefVal = r[refCol]?.ToString().Trim() ?? "";
                                
                                if (filterVal == "非空值 (有輸入即算)" || string.IsNullOrEmpty(filterVal)) {
                                    if (string.IsNullOrEmpty(rowRefVal)) return false;
                                } else {
                                    bool isMatch = false;
                                    var cellValues = rowRefVal.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var v in cellValues) {
                                        if (v.Trim().Equals(filterVal, StringComparison.OrdinalIgnoreCase)) {
                                            isMatch = true; break;
                                        }
                                    }
                                    if (!isMatch) return false;
                                }
                            }

                            return true;
                        }).ToList();

                        if (agg == "COUNT") {
                            if (fCol == "Id (無條件計數)" || fCol == "Id" || !fDt.Columns.Contains(fCol)) {
                                computedVal = matchedRows.Count;
                            } else {
                                computedVal = matchedRows.Count(r => r[fCol] != DBNull.Value && !string.IsNullOrWhiteSpace(r[fCol].ToString()));
                            }
                        }
                        else if (fDt.Columns.Contains(fCol)) {
                            List<double> vals = new List<double>();
                            foreach (var fr in matchedRows) {
                                if (double.TryParse(fr[fCol]?.ToString().Replace(",", ""), out double v)) vals.Add(v);
                            }

                            if (vals.Count > 0) {
                                if (agg == "SUM") computedVal = vals.Sum();
                                else if (agg == "AVG") computedVal = vals.Average();
                                else if (agg == "MAX") computedVal = vals.Max();
                                else if (agg == "MIN") computedVal = vals.Min();
                            }
                        }
                    }
                    expr = expr.Replace(mCross.Value, computedVal.ToString());
                }

                try {
                    DataTable dtMath = new DataTable();
                    object result = dtMath.Compute(expr, null);
                    if (result != DBNull.Value) {
                        double dRes = Convert.ToDouble(result);
                        evalText = evalText.Replace(mBlock.Value, Math.Round(dRes, 2).ToString("0.##"));
                    }
                } catch {
                    evalText = evalText.Replace(mBlock.Value, "NaN");
                }
            }
            return evalText;
        }

        private async Task CalculateAndLoadGrid(ThemeSectionUI ui)
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            string startYM = $"{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}";
            string endYM = $"{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}";

            ui.Dgv.Rows.Clear();

            await Task.Run(() => {
                DataTable dtConfigs = new DataTable();
                DataTable dtRecords = new DataTable();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblConfigs} WHERE ThemeId=@T", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtConfigs);
                        }
                        
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblRecords} WHERE ThemeId=@T", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtRecords);
                        }
                    }
                } catch { return; }

                Dictionary<string, (HashSet<string> Atts, HashSet<string> Rems)> mergedHistory = new Dictionary<string, (HashSet<string>, HashSet<string>)>();

                foreach (DataRow r in dtRecords.Rows) {
                    string recS = r["PeriodStart"].ToString();
                    string recE = r["PeriodEnd"].ToString();
                    string itemName = r["ItemName"].ToString();

                    if (string.Compare(recS, endYM) <= 0 && string.Compare(recE, startYM) >= 0) {
                        
                        if (!mergedHistory.ContainsKey(itemName)) {
                            mergedHistory[itemName] = (new HashSet<string>(), new HashSet<string>());
                        }

                        string att = r["Attachment"]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(att)) {
                            foreach (string p in att.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)) {
                                mergedHistory[itemName].Atts.Add(p.Trim());
                            }
                        }

                        string rem = r["Remarks"]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(rem)) {
                            mergedHistory[itemName].Rems.Add(rem.Trim());
                        }
                    }
                }

                Dictionary<string, DataTable> tableCache = new Dictionary<string, DataTable>();

                foreach (DataRow cfgRow in dtConfigs.Rows)
                {
                    string itemName = cfgRow["ItemName"].ToString();
                    string template = cfgRow["FormulaTemplate"].ToString();

                    string evalText = EvaluateStatsFormula(template, startYM, endYM, tableCache);

                    string finalAttach = "";
                    string finalRemarks = "";
                    if (mergedHistory.ContainsKey(itemName)) {
                        finalAttach = string.Join("|", mergedHistory[itemName].Atts);
                        finalRemarks = string.Join("\r\n\r\n", mergedHistory[itemName].Rems);
                    }

                    if (ui.Dgv.InvokeRequired) {
                        ui.Dgv.Invoke(new Action(() => ui.Dgv.Rows.Add(itemName, evalText, finalAttach, finalRemarks)));
                    } else {
                        ui.Dgv.Rows.Add(itemName, evalText, finalAttach, finalRemarks);
                    }
                }
            });

            if (ui.Dgv.InvokeRequired) {
                ui.Dgv.Invoke(new Action(() => {
                    ApplyColumnWidths(ui);
                    AdjustGridHeight(ui);
                }));
            } else {
                ApplyColumnWidths(ui);
                AdjustGridHeight(ui);
            }

            ui.MainBox.Invalidate(); 
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private async Task RecalculateGridData(ThemeSectionUI ui)
        {
            if (ui.Dgv.Rows.Count == 0) {
                MessageBox.Show("目前表格沒有資料，請先點擊「🔍 讀取」。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
            ui.Dgv.EndEdit(); 

            string startYM = $"{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}";
            string endYM = $"{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}";

            Dictionary<string, string> newValues = new Dictionary<string, string>();

            await Task.Run(() => {
                DataTable dtConfigs = new DataTable();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn))
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtConfigs);
                    }
                } catch { return; }

                Dictionary<string, DataTable> tableCache = new Dictionary<string, DataTable>();

                foreach (DataRow cfgRow in dtConfigs.Rows)
                {
                    string itemName = cfgRow["ItemName"].ToString();
                    string template = cfgRow["FormulaTemplate"].ToString();
                    newValues[itemName] = EvaluateStatsFormula(template, startYM, endYM, tableCache);
                }
            });

            Action updateGrid = () => {
                foreach (DataGridViewRow row in ui.Dgv.Rows) {
                    if (row.IsNewRow) continue;
                    string itemName = row.Cells["項目"].Value?.ToString() ?? "";
                    if (newValues.ContainsKey(itemName)) {
                        row.Cells["數據"].Value = newValues[itemName];
                    }
                }
                AdjustGridHeight(ui);
            };

            if (ui.Dgv.InvokeRequired) ui.Dgv.Invoke(updateGrid);
            else updateGrid();

            ui.MainBox.Invalidate();
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            
            MessageBox.Show("重新統計完成！\n\n最新的運算數值已更新至「數據」欄位，您畫面上的「附件與備註」均保持不變。\n若確認無誤，請點擊「💾 儲存」將最終結果寫入資料庫。", "統計更新完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SaveGridData(ThemeSectionUI ui)
        {
            ui.Dgv.EndEdit();
            string startYM = $"{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}";
            string endYM = $"{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}";

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        
                        new SQLiteCommand($"DELETE FROM {TblRecords} WHERE ThemeId={ui.ThemeId} AND PeriodStart='{startYM}' AND PeriodEnd='{endYM}'", conn, trans).ExecuteNonQuery();

                        string sql = $"INSERT INTO {TblRecords} (ThemeId, PeriodStart, PeriodEnd, ItemName, DataValue, Attachment, Remarks) VALUES (@T, @PS, @PE, @IN, @DV, @A, @R)";
                        foreach (DataGridViewRow row in ui.Dgv.Rows) {
                            if (row.IsNewRow) continue;
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                                cmd.Parameters.AddWithValue("@PS", startYM);
                                cmd.Parameters.AddWithValue("@PE", endYM);
                                cmd.Parameters.AddWithValue("@IN", row.Cells["項目"].Value?.ToString() ?? "");
                                cmd.Parameters.AddWithValue("@DV", row.Cells["數據"].Value?.ToString() ?? "");
                                cmd.Parameters.AddWithValue("@A", row.Cells["附件檔案"].Value?.ToString() ?? "");
                                cmd.Parameters.AddWithValue("@R", row.Cells["備註"].Value?.ToString() ?? "");
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                MessageBox.Show($"主題【{ui.ThemeName}】 區間 {startYM} ~ {endYM} 的資料儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                MessageBox.Show("儲存失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && dgv != null) {
                if (dgv.Columns[e.ColumnIndex].Name == "附件檔案" && e.Value != null) {
                    string pathStr = e.Value.ToString();
                    if (!string.IsNullOrEmpty(pathStr)) {
                        string[] parts = pathStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 1) { e.Value = $"📁 [共 {parts.Length} 個檔案]"; } 
                        else { e.Value = Path.GetFileName(parts[0]); }
                        e.FormattingApplied = true;
                    }
                }
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e, ThemeSectionUI ui)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && !ui.Dgv.Rows[e.RowIndex].IsNewRow) 
            {
                if (ui.Dgv.Columns[e.ColumnIndex].Name == "附件檔案") 
                {
                    string currentVal = ui.Dgv[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
                    string targetFolder = $"{ui.ThemeId}/{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}_{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}"; 
                    
                    using (var frm = new AttachmentManagerUI(currentVal, "Menu3Stats", ui.ThemeId.ToString(), targetFolder, delegate(string path) { DeletePhysicalFile(path, ui.Dgv, e.RowIndex); })) {
                        if (frm.ShowDialog() == DialogResult.OK) { 
                            ui.Dgv[e.ColumnIndex, e.RowIndex].Value = frm.FinalPathsString; 
                            ui.Dgv.EndEdit(); 
                        }
                    }
                }
            }
        }

        private void DeletePhysicalFile(string relativePath, DataGridView dgv, int currentRowIndex) 
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;
            bool isUsedByOthers = false;
            
            foreach (DataGridViewRow row in dgv.Rows) {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                object cellVal = row.Cells["附件檔案"].Value;
                if (cellVal != null) { 
                    string[] paths = cellVal.ToString().Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries); 
                    foreach(string p in paths) { if (p == relativePath) { isUsedByOthers = true; break; } }
                }
            }
            
            if (isUsedByOthers) return;
            
            try {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) File.Delete(absPath); 
            } catch { }
        }

        private void OpenSettingsDialog(ThemeSectionUI ui)
        {
            using (Form f = new Form { Text = $"⚙️ 設定顯示與公式 - {ui.ThemeName}", Size = new Size(1380, 780), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                
                Button btnAddNew = new Button { Text = "➕ 新增統計項目", Dock = DockStyle.Top, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
                btnAddNew.FlatAppearance.BorderSize = 0;

                Label l1 = new Label { Text = "清單項目 (拖曳可排序)", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 35, Padding = new Padding(0,10,0,0) };
                
                DataGridView dgvItems = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows=false, RowHeadersVisible=false, ColumnHeadersVisible=false, SelectionMode=DataGridViewSelectionMode.FullRowSelect, BackgroundColor=Color.White, AllowDrop=true, MultiSelect=false };
                dgvItems.Columns.Add("Name", "Name"); dgvItems.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; dgvItems.Columns["Name"].ReadOnly = true;
                
                typeof(DataGridView).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, null, dgvItems, new object[] { true });

                Panel pnlLeftActions = new Panel { Dock = DockStyle.Bottom, Height = 120, Padding = new Padding(0, 10, 0, 0) };
                Button btnDown = new Button { Text = "↓ 下移", Dock = DockStyle.Top, Height = 35, BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Panel spacer1 = new Panel { Dock = DockStyle.Top, Height = 5 }; 
                Button btnUp = new Button { Text = "↑ 上移", Dock = DockStyle.Top, Height = 35, BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };

                pnlLeftActions.Controls.Add(btnDown);
                pnlLeftActions.Controls.Add(spacer1);
                pnlLeftActions.Controls.Add(btnUp);
                pnlLeftActions.Controls.Add(btnDel);

                pnlLeft.Controls.Add(dgvItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnAddNew);
                pnlLeft.Controls.Add(pnlLeftActions);

                // ============== 右側編輯區 ==============
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯選取項目的內容公式", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                Panel pName = new Panel { Width = 1000, Height = 45 };
                pName.Controls.Add(new Label { Text = "項目名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                // 🟢 調整 1：項目名稱文字框加寬 100px (350 -> 450)
                TextBox txtName = new TextBox { Width = 450, Location = new Point(120, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);
                flpEditor.Controls.Add(pName);

                // 🟢 變數產生器
                GroupBox boxBuilder = new GroupBox { Text = "變數產生器 (自動產生跨表聚合公式)", Width = 1000, Height = 135, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding = new Padding(10) };
                
                Panel pnlBuilderInner = new Panel { Dock = DockStyle.Fill };
                
                // 第一排：庫、表、被計算欄、日期欄 (Y=10)
                ComboBox cbDb = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbTb = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbCol = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                cbCol.Items.Add("Id (無條件計數)");
                ComboBox cbDateCol = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };

                Label lblDb = new Label { Text = "庫:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                Label lblTb = new Label { Text = "表:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                Label lblCol = new Label { Text = "被計算欄:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                Label lblDateCol = new Label { Text = "日期欄:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };

                // 🟢 調整 2 & 3：第一排座標重新分配，下拉選單間距+20
                lblDb.Location = new Point(5, 13);
                cbDb.Location = new Point(40, 10);
                
                lblTb.Location = new Point(190, 13);
                cbTb.Location = new Point(225, 10);
                
                lblCol.Location = new Point(415, 13);
                cbCol.Location = new Point(510, 10); // 增加 20 間距
                
                lblDateCol.Location = new Point(700, 13);
                cbDateCol.Location = new Point(780, 10); // 增加 20 間距

                pnlBuilderInner.Controls.AddRange(new Control[] { lblDb, cbDb, lblTb, cbTb, lblCol, cbCol, lblDateCol, cbDateCol });

                // 第二排：篩選條件欄、內容、動作、插入按鈕 (Y=60)
                ComboBox cbRefCol = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbFilterVal = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                
                ComboBox cbAction = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                cbAction.Items.AddRange(new string[] { "加總 (SUM)", "平均值 (AVG)", "最大值 (MAX)", "最小值 (MIN)", "計數 (COUNT)" }); 
                cbAction.SelectedIndex = 0;

                Button btnInsert = new Button { Text = "插入變數 ⬇️", Width = 120, Height = 34, BackColor = Color.DarkCyan, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnInsert.FlatAppearance.BorderSize = 0;

                Label lblRefCol = new Label { Text = "篩選條件欄:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                Label lblFilterVal = new Label { Text = "內容:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                Label lblAction = new Label { Text = "動作:", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };

                // 🟢 調整 4, 5, 6：第二排座標重新分配，下拉選單間距+20
                lblRefCol.Location = new Point(5, 63);
                cbRefCol.Location = new Point(115, 60); // 增加 20 間距

                lblFilterVal.Location = new Point(275, 63);
                cbFilterVal.Location = new Point(335, 60); // 增加 20 間距

                lblAction.Location = new Point(505, 63);
                cbAction.Location = new Point(565, 60); // 增加 20 間距

                btnInsert.Location = new Point(720, 58); // 按鈕往右推

                pnlBuilderInner.Controls.AddRange(new Control[] { lblRefCol, cbRefCol, lblFilterVal, cbFilterVal, lblAction, cbAction, btnInsert });

                boxBuilder.Controls.Add(pnlBuilderInner);

                // 綁定中英文對照
                cbDb.Items.Add(new ItemMap { EnName="", ChName="" });
                foreach(var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName=kvp.Key, ChName=kvp.Value.ChDbName });

                cbDb.SelectedIndexChanged += (s, e) => {
                    cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName="", ChName="" });
                    var db = cbDb.SelectedItem as ItemMap;
                    if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                        foreach(var tb in _dbMap[db.EnName].Tables) cbTb.Items.Add(new ItemMap{ EnName=tb.Key, ChName=tb.Value });
                    }
                };

                cbTb.SelectedIndexChanged += (s, e) => {
                    cbCol.Items.Clear(); cbCol.Items.Add("Id (無條件計數)");
                    cbDateCol.Items.Clear(); cbDateCol.Items.Add("自動判斷");
                    cbRefCol.Items.Clear(); cbRefCol.Items.Add("");
                    
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                        var cols = DataManager.GetColumnNames(db.EnName, tb.EnName);
                        
                        foreach(var c in cols.Where(x => x != "Id" && !x.Contains("日期") && !x.Contains("年月"))) {
                            cbCol.Items.Add(c);
                        }

                        foreach(var c in cols) {
                            cbDateCol.Items.Add(c);
                            if (c != "Id") cbRefCol.Items.Add(c);
                        }

                        if (cols.Contains("日期")) cbDateCol.SelectedItem = "日期";
                        else if (cols.Contains("年月")) cbDateCol.SelectedItem = "年月";
                        else if (cols.Contains("清運日期")) cbDateCol.SelectedItem = "清運日期";
                        else if (cols.Contains("年度")) cbDateCol.SelectedItem = "年度";
                        else cbDateCol.SelectedIndex = 0;
                    }
                };

                cbRefCol.SelectedIndexChanged += (s, e) => {
                    cbFilterVal.Items.Clear();
                    cbFilterVal.Items.Add("非空值 (有輸入即算)");
                    if (cbDb.SelectedItem != null && cbTb.SelectedItem != null && !string.IsNullOrEmpty(cbRefCol.Text)) {
                        string db = ((ItemMap)cbDb.SelectedItem).EnName;
                        string tb = ((ItemMap)cbTb.SelectedItem).EnName;
                        string col = cbRefCol.Text;
                        try {
                            DataTable dt = DataManager.GetTableData(db, tb, "", "", "");
                            if (dt != null && dt.Columns.Contains(col)) {
                                var distinctVals = dt.Rows.Cast<DataRow>()
                                    .Select(r => r[col]?.ToString().Trim())
                                    .Where(str => !string.IsNullOrEmpty(str))
                                    .Distinct();
                                foreach (var v in distinctVals) cbFilterVal.Items.Add(v);
                            }
                        } catch { }
                    }
                    cbFilterVal.SelectedIndex = 0; 
                };

                cbAction.SelectedIndexChanged += (s, e) => {
                    if (cbAction.Text.Contains("COUNT") && cbCol.Items.Contains("Id (無條件計數)")) {
                        cbCol.SelectedItem = "Id (無條件計數)";
                    }
                };

                flpEditor.Controls.Add(boxBuilder);

                Label lblDesc = new Label { 
                    Text = "混合圖文公式編輯區：\n(支援多行排版，請直接按 Enter 換行。純文字打在外面，需計算的公式包在 { 大括號 } 內)", 
                    AutoSize = true, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Margin = new Padding(0,10,0,5), 
                    ForeColor=Color.DarkMagenta 
                };
                
                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Width=1000, Height = 40, WrapContents = false };
                string[] keys = { "+", "-", "*", "/", "(", ")", "{", "}" };
                
                RichTextBox rtbFormula = new RichTextBox { 
                    Width = 1000, 
                    Height = 205, 
                    Font = new Font("Consolas", 14F), 
                    BackColor = Color.AliceBlue, 
                    Margin = new Padding(0, 5, 0, 0),
                    AcceptsTab = true 
                };
                
                rtbFormula.KeyDown += (s, ev) => {
                    if (ev.KeyCode == Keys.Enter) {
                        int selectionStart = rtbFormula.SelectionStart;
                        rtbFormula.Text = rtbFormula.Text.Insert(selectionStart, Environment.NewLine);
                        rtbFormula.SelectionStart = selectionStart + Environment.NewLine.Length;
                        ev.Handled = true;
                    }
                };
                
                foreach (var k in keys) {
                    Button b = new Button { Text = k, Width = 45, Height = 35, Font=new Font("Consolas", 14F, FontStyle.Bold), Cursor=Cursors.Hand, BackColor=Color.WhiteSmoke };
                    pnlKeys.Controls.Add(b);
                    b.Click += (s, e) => { rtbFormula.Focus(); rtbFormula.SelectedText = $" {b.Text} "; };
                }

                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇完整的跨表欄位來源！"); return; }
                    
                    string actionText = cbAction.Text;
                    string aggCode = "SUM";
                    if (actionText.Contains("SUM")) aggCode = "SUM";
                    else if (actionText.Contains("AVG")) aggCode = "AVG";
                    else if (actionText.Contains("MAX")) aggCode = "MAX";
                    else if (actionText.Contains("MIN")) aggCode = "MIN";
                    else if (actionText.Contains("COUNT")) aggCode = "COUNT";

                    string baseStr = $"{aggCode}([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}]";
                    string dateColStr = (cbDateCol.SelectedItem != null && cbDateCol.SelectedItem.ToString() != "自動判斷") ? $".[{(cbDateCol.SelectedItem)}]" : ".[自動判斷]";
                    
                    if (!string.IsNullOrEmpty(cbRefCol.Text)) {
                        string filterStr = cbFilterVal.Text.Trim();
                        if (string.IsNullOrEmpty(filterStr)) filterStr = "非空值 (有輸入即算)";
                        baseStr += $"{dateColStr}.[{(cbRefCol.Text)}].[{filterStr}]";
                    } else if (dateColStr != ".[自動判斷]") {
                        baseStr += dateColStr;
                    }

                    baseStr += ")";
                    rtbFormula.Focus();
                    rtbFormula.SelectedText = baseStr;
                };

                flpEditor.Controls.Add(lblDesc);
                flpEditor.Controls.Add(pnlKeys);
                flpEditor.Controls.Add(rtbFormula);

                Button btnSaveRow = new Button { Text = "💾 儲存並加入清單", Width = 900, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 0), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSaveRow.FlatAppearance.BorderSize = 0;

                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);
                pnlRight.Controls.Add(btnSaveRow);
                btnSaveRow.Dock = DockStyle.Bottom;

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);
                f.Controls.Add(tlp);

                var configs = new List<EditingConfig>();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT ItemName, FormulaTemplate FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn))
                        using (var reader = cmd.ExecuteReader()) {
                            while(reader.Read()) {
                                configs.Add(new EditingConfig { Name = reader[0].ToString(), Formula = reader[1].ToString() });
                            }
                        }
                    }
                } catch { }

                bool isSyncing = false;

                Action refreshList = () => { 
                    isSyncing = true;
                    dgvItems.Rows.Clear(); 
                    foreach(var c in configs) dgvItems.Rows.Add(c.Name); 
                    isSyncing = false;
                };

                int dragFromIdx = -1;
                int dragToIdx = -1;
                Rectangle dragBox = Rectangle.Empty;

                dgvItems.MouseDown += (s, e) => {
                    var hit = dgvItems.HitTest(e.X, e.Y);
                    if (hit.RowIndex >= 0) {
                        dragFromIdx = hit.RowIndex;
                        Size dragSize = SystemInformation.DragSize;
                        dragBox = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
                    } else {
                        dragBox = Rectangle.Empty;
                    }
                };

                dgvItems.MouseMove += (s, e) => {
                    if ((e.Button & MouseButtons.Left) == MouseButtons.Left) {
                        if (dragBox != Rectangle.Empty && !dragBox.Contains(e.X, e.Y)) {
                            dgvItems.DoDragDrop(dgvItems.Rows[dragFromIdx], DragDropEffects.Move);
                        }
                    }
                };

                dgvItems.DragOver += (s, e) => {
                    e.Effect = DragDropEffects.Move;
                    Point p = dgvItems.PointToClient(new Point(e.X, e.Y));
                    var hit = dgvItems.HitTest(p.X, p.Y);
                    int newToIdx = hit.RowIndex;
                    if (newToIdx < 0) newToIdx = dgvItems.Rows.Count - 1;

                    if (dragToIdx != newToIdx) {
                        dragToIdx = newToIdx;
                        dgvItems.Invalidate(); 
                    }
                };

                dgvItems.DragDrop += (s, e) => {
                    Point p = dgvItems.PointToClient(new Point(e.X, e.Y));
                    var hit = dgvItems.HitTest(p.X, p.Y);
                    int targetIdx = hit.RowIndex;
                    if (targetIdx < 0) targetIdx = dgvItems.Rows.Count - 1;

                    if (dragFromIdx >= 0 && dragFromIdx != targetIdx) {
                        var item = configs[dragFromIdx];
                        configs.RemoveAt(dragFromIdx);
                        configs.Insert(targetIdx, item);

                        refreshList(); 
                        dgvItems.ClearSelection();
                        dgvItems.Rows[targetIdx].Selected = true;
                    }
                    dragToIdx = -1;
                    dgvItems.Invalidate();
                };

                dgvItems.DragLeave += (s, e) => {
                    dragToIdx = -1;
                    dgvItems.Invalidate();
                };

                dgvItems.Paint += (s, e) => {
                    if (dragToIdx >= 0 && dragToIdx < dgvItems.Rows.Count) {
                        Rectangle r = dgvItems.GetRowDisplayRectangle(dragToIdx, false);
                        using (Pen pen = new Pen(Color.Red, 3)) {
                            e.Graphics.DrawLine(pen, r.Left, r.Top, r.Right, r.Top);
                        }
                    }
                };

                btnUp.Click += (s, e) => {
                    if (dgvItems.SelectedRows.Count > 0) {
                        int idx = dgvItems.SelectedRows[0].Index;
                        if (idx > 0) {
                            var item = configs[idx];
                            configs.RemoveAt(idx);
                            configs.Insert(idx - 1, item);
                            
                            refreshList();
                            dgvItems.ClearSelection();
                            dgvItems.Rows[idx - 1].Selected = true;
                        }
                    }
                };

                btnDown.Click += (s, e) => {
                    if (dgvItems.SelectedRows.Count > 0) {
                        int idx = dgvItems.SelectedRows[0].Index;
                        if (idx >= 0 && idx < configs.Count - 1) {
                            var item = configs[idx];
                            configs.RemoveAt(idx);
                            configs.Insert(idx + 1, item);
                            
                            refreshList();
                            dgvItems.ClearSelection();
                            dgvItems.Rows[idx + 1].Selected = true;
                        }
                    }
                };

                dgvItems.SelectionChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    isSyncing = true;
                    txtName.Text = configs[idx].Name;
                    rtbFormula.Text = configs[idx].Formula;
                    isSyncing = false;
                };

                txtName.TextChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    configs[idx].Name = txtName.Text;
                    
                    isSyncing = true; 
                    dgvItems.Rows[idx].Cells[0].Value = txtName.Text; 
                    isSyncing = false;
                };

                rtbFormula.TextChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    configs[idx].Formula = rtbFormula.Text;
                };

                btnAddNew.Click += (s, e) => {
                    configs.Add(new EditingConfig { Name = "新項目", Formula = "" }); 
                    refreshList();
                    dgvItems.ClearSelection(); 
                    dgvItems.Rows[dgvItems.Rows.Count-1].Selected = true;
                    txtName.Focus();
                };

                btnDel.Click += (s, e) => {
                    if (dgvItems.SelectedRows.Count > 0) {
                        int idx = dgvItems.SelectedRows[0].Index;
                        configs.RemoveAt(idx); 
                        refreshList();
                        
                        if (configs.Count > 0) {
                            dgvItems.Rows[Math.Min(idx, configs.Count - 1)].Selected = true;
                        } else {
                            txtName.Clear(); rtbFormula.Clear();
                        }
                    }
                };

                btnSaveRow.Click += (s, e) => {
                    btnSaveRow.Focus(); 
                    
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                new SQLiteCommand($"DELETE FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn, trans).ExecuteNonQuery();
                                
                                string sql = $"INSERT INTO {TblConfigs} (ThemeId, ItemName, FormulaTemplate) VALUES ({ui.ThemeId}, @N, @F)";
                                foreach(DataGridViewRow dgvRow in dgvItems.Rows) {
                                    string currentName = dgvRow.Cells[0].Value.ToString();
                                    var targetConfig = configs.FirstOrDefault(c => c.Name == currentName);
                                    
                                    if(targetConfig == null || string.IsNullOrWhiteSpace(targetConfig.Name)) continue;
                                    
                                    using(var cmd = new SQLiteCommand(sql, conn, trans)) {
                                        cmd.Parameters.AddWithValue("@N", targetConfig.Name); 
                                        cmd.Parameters.AddWithValue("@F", targetConfig.Formula);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                trans.Commit();
                            }
                        }
                        f.DialogResult = DialogResult.OK;
                    } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
                };

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(10) };
                Button btnExp = new Button { Text = "📤 匯出設定", Dock = DockStyle.Left, Width = 130, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnImp = new Button { Text = "📥 匯入設定", Dock = DockStyle.Left, Width = 130, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                btnExp.Click += (s, e) => {
                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = $"統計設定_{ui.ThemeName}_{DateTime.Now:yyyyMMdd}" }) {
                        if (sfd.ShowDialog() == DialogResult.OK) {
                            try {
                                DataTable dt = new DataTable(); dt.Columns.Add("顯示名稱"); dt.Columns.Add("文字與運算公式組合");
                                foreach(var c in configs) dt.Rows.Add(c.Name, c.Formula);
                                using (ExcelPackage p = new ExcelPackage()) {
                                    var ws = p.Workbook.Worksheets.Add("設定"); ws.Cells["A1"].LoadFromDataTable(dt, true); ws.Cells.AutoFitColumns();
                                    p.SaveAs(new FileInfo(sfd.FileName));
                                }
                                MessageBox.Show("匯出成功！");
                            } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message); }
                        }
                    }
                };

                btnImp.Click += (s, e) => {
                    using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) {
                        if (ofd.ShowDialog() == DialogResult.OK) {
                            try {
                                using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                                    ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                    if (ws == null || ws.Dimension == null) return;
                                    configs.Clear();
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string n = ws.Cells[r, 1].Text.Trim(); string fma = ws.Cells[r, 2].Text.Trim();
                                        if(!string.IsNullOrEmpty(n)) configs.Add(new EditingConfig { Name = n, Formula = fma });
                                    }
                                }
                                refreshList();
                                if (configs.Count > 0) dgvItems.Rows[0].Selected = true;
                                MessageBox.Show("匯入成功！請點擊右下角「儲存全部設定」以生效。");
                            } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message); }
                        }
                    }
                };

                pnlBottom.Controls.Add(btnImp); pnlBottom.Controls.Add(btnExp);
                f.Controls.Add(pnlBottom);

                refreshList();
                if (dgvItems.Rows.Count > 0) dgvItems.Rows[0].Selected = true;
                
                f.ShowDialog();
            }
        }

        // ==========================================
        // 🟢 全域 PDF / Excel 導出對話框與邏輯
        // ==========================================
        private List<ThemeSectionUI> GetSelectedThemesDialog()
        {
            List<ThemeSectionUI> selected = new List<ThemeSectionUI>();
            if (_sections.Count == 0) { MessageBox.Show("目前沒有任何統計區塊可供匯出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return selected; }

            using (Form f = new Form() { Width = 450, Height = 500, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));

                Label lbl = new Label { Text = "請勾選欲匯出的統計主題區塊：", Dock = DockStyle.Fill, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), AutoSize = true };
                tlp.Controls.Add(lbl, 0, 0);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(15, 5, 15, 5), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
                
                foreach (var sec in _sections) {
                    clb.Items.Add(sec.ThemeName, true); 
                }
                tlp.Controls.Add(clb, 0, 1);

                Panel pnlBottom = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) };
                
                Button btnSelectAll = new Button { Text = "☑️ 全選", Location = new Point(15, 5), Size = new Size(100, 35), BackColor = Color.LightGray, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button btnUnselectAll = new Button { Text = "☐ 取消全選", Location = new Point(125, 5), Size = new Size(130, 35), BackColor = Color.LightGray, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 40, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                
                btnSelectAll.Click += (s, ev) => {
                    for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true);
                };

                btnUnselectAll.Click += (s, ev) => {
                    for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false);
                };

                pnlBottom.Controls.Add(btnSelectAll);
                pnlBottom.Controls.Add(btnUnselectAll);
                pnlBottom.Controls.Add(btnOk);
                
                tlp.Controls.Add(pnlBottom, 0, 2);
                f.Controls.Add(tlp);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    for (int i = 0; i < clb.Items.Count; i++) {
                        if (clb.GetItemChecked(i)) {
                            selected.Add(_sections[i]);
                        }
                    }
                }
            }
            return selected;
        }

        private void BtnGlobalPdf_Click(object sender, EventArgs e)
        {
            var selectedSections = GetSelectedThemesDialog();
            if (selectedSections.Count == 0) return;

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = $"選單3_動態看板報表_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                    try 
                    {
                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = true;
                        pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                        int currentSectionIdx = 0;
                        int currentRowIdx = 0;
                        int pageNumber = 1;
                        bool drawSectionHeader = true;

                        pd.PrintPage += (s, ev) => 
                        {
                            Graphics g = ev.Graphics;
                            float x = ev.MarginBounds.Left;
                            float y = ev.MarginBounds.Top;
                            float w = ev.MarginBounds.Width;

                            Font fMainTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                            Font fSubTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                            Font fSign = new Font("Microsoft JhengHei UI", 12F);
                            Font fTheme = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                            Font fPeriod = new Font("Microsoft JhengHei UI", 11F);
                            Font fHead = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                            Font fBody = new Font("Microsoft JhengHei UI", 10F);
                            Font fFooter = new Font("Microsoft JhengHei UI", 10F);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                            StringFormat sfWrap = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fMainTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                            y += 40;
                            g.DrawString("動態統計看板綜合報表", fSubTitle, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                            y += 40;
                            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
                            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                            y += 40;

                            while (currentSectionIdx < selectedSections.Count)
                            {
                                var sec = selectedSections[currentSectionIdx];
                                var dgv = sec.Dgv;

                                if (drawSectionHeader)
                                {
                                    if (y + 120 > ev.MarginBounds.Bottom) { 
                                        g.DrawString($"- {pageNumber} -", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, w, 20), sfCenter);
                                        pageNumber++;
                                        ev.HasMorePages = true; 
                                        return; 
                                    }

                                    g.DrawString($"■ {sec.ThemeName}", fTheme, Brushes.DarkSlateBlue, x, y);
                                    y += 30;

                                    string startYM = $"{sec.CboStartYear.Text}年{sec.CboStartMonth.Text}月";
                                    string endYM = $"{sec.CboEndYear.Text}年{sec.CboEndMonth.Text}月";
                                    g.DrawString($"查詢區間: {startYM} ~ {endYM}", fPeriod, Brushes.DimGray, x + 10, y);
                                    y += 30;

                                    float currX = x;
                                    float[] colWidths = new float[dgv.Columns.Count];
                                    float totalGridWidth = dgv.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);

                                    for (int i = 0; i < dgv.Columns.Count; i++) {
                                        colWidths[i] = (dgv.Columns[i].Width / totalGridWidth) * w;
                                        RectangleF rHead = new RectangleF(currX, y, colWidths[i], 35);
                                        g.FillRectangle(Brushes.LightGray, rHead);
                                        g.DrawRectangle(Pens.Black, Rectangle.Round(rHead));
                                        g.DrawString(dgv.Columns[i].HeaderText, fHead, Brushes.Black, rHead, sfCenter);
                                        currX += colWidths[i];
                                    }
                                    y += 35;
                                    drawSectionHeader = false;
                                }

                                float[] cW = new float[dgv.Columns.Count];
                                float tW = dgv.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                                for (int i = 0; i < dgv.Columns.Count; i++) cW[i] = (dgv.Columns[i].Width / tW) * w;

                                while (currentRowIdx < dgv.Rows.Count)
                                {
                                    DataGridViewRow row = dgv.Rows[currentRowIdx];
                                    float maxH = 35;
                                    
                                    for (int i = 0; i < dgv.Columns.Count; i++) {
                                        string val = row.Cells[i].Value?.ToString() ?? "";
                                        SizeF sz = g.MeasureString(val, fBody, (int)cW[i], sfWrap);
                                        if (sz.Height + 10 > maxH) maxH = sz.Height + 10;
                                    }

                                    if (y + maxH > ev.MarginBounds.Bottom - 30) {
                                        g.DrawString($"- {pageNumber} -", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, w, 20), sfCenter);
                                        pageNumber++;
                                        ev.HasMorePages = true;
                                        return;
                                    }

                                    float rX = x;
                                    for (int i = 0; i < dgv.Columns.Count; i++) {
                                        RectangleF rCell = new RectangleF(rX, y, cW[i], maxH);
                                        g.DrawRectangle(Pens.Black, Rectangle.Round(rCell));
                                        string val = row.Cells[i].Value?.ToString() ?? "";
                                        RectangleF textRect = new RectangleF(rCell.X + 2, rCell.Y + 2, rCell.Width - 4, rCell.Height - 4);
                                        g.DrawString(val, fBody, Brushes.Black, textRect, sfWrap);
                                        rX += cW[i];
                                    }

                                    y += maxH;
                                    currentRowIdx++;
                                }

                                y += 30; 
                                currentSectionIdx++;
                                currentRowIdx = 0;
                                drawSectionHeader = true;
                            }

                            g.DrawString($"- {pageNumber} -", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, w, 20), sfCenter);
                            ev.HasMorePages = false;
                        };

                        pd.Print();
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void BtnGlobalExcel_Click(object sender, EventArgs e)
        {
            var selectedSections = GetSelectedThemesDialog();
            if (selectedSections.Count == 0) return;

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = $"選單3_動態看板報表_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        using (ExcelPackage p = new ExcelPackage())
                        {
                            foreach (var sec in selectedSections)
                            {
                                if (sec.Dgv.Rows.Count == 0) continue;

                                DataTable dt = new DataTable();
                                foreach (DataGridViewColumn col in sec.Dgv.Columns) dt.Columns.Add(col.HeaderText.Replace("\n", ""));
                                
                                foreach (DataGridViewRow row in sec.Dgv.Rows) {
                                    if (row.IsNewRow) continue;
                                    DataRow dRow = dt.NewRow();
                                    for (int i = 0; i < sec.Dgv.Columns.Count; i++) {
                                        dRow[i] = row.Cells[i].Value?.ToString() ?? "";
                                    }
                                    dt.Rows.Add(dRow);
                                }

                                string safeSheetName = sec.ThemeName;
                                foreach (char c in new[] { '*', ':', '?', '[', ']', '\\', '/' }) {
                                    safeSheetName = safeSheetName.Replace(c.ToString(), "");
                                }
                                if (safeSheetName.Length > 31) safeSheetName = safeSheetName.Substring(0, 31);
                                
                                int duplicateCount = 1;
                                string finalSheetName = safeSheetName;
                                while(p.Workbook.Worksheets.Any(sheet => sheet.Name == finalSheetName)) {
                                    finalSheetName = $"{safeSheetName}_{duplicateCount}";
                                    duplicateCount++;
                                }

                                var ws = p.Workbook.Worksheets.Add(finalSheetName);
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
                                
                                float totalGridWidth = sec.Dgv.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                                for (int i = 0; i < sec.Dgv.Columns.Count; i++) {
                                    float excelWidth = (sec.Dgv.Columns[i].Width / totalGridWidth) * 120f;
                                    if (excelWidth < 10) excelWidth = 10;
                                    ws.Column(i + 1).Width = excelWidth;
                                }

                                using (var range = ws.Cells[1, 1, 1, dt.Columns.Count]) {
                                    range.Style.Font.Bold = true;
                                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                }

                                var dataRange = ws.Cells[2, 1, dt.Rows.Count + 1, dt.Columns.Count];
                                dataRange.Style.WrapText = true;
                                dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            }
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("Excel 綜合報表匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private string ShowInputBox(string prompt, string title, string defaultValue)
        {
            using (Form form = new Form { Width = 400, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = title, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White })
            {
                Label label = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox textBox = new TextBox() { Left = 20, Top = 60, Width = 340, Text = defaultValue, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button confirmation = new Button() { Text = "確認", Left = 160, Width = 90, Height = 35, Top = 100, DialogResult = DialogResult.OK, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button cancel = new Button() { Text = "取消", Left = 270, Width = 90, Height = 35, Top = 100, DialogResult = DialogResult.Cancel, Font = new Font("Microsoft JhengHei UI", 11F) };

                form.Controls.Add(label); form.Controls.Add(textBox); form.Controls.Add(confirmation); form.Controls.Add(cancel);
                form.AcceptButton = confirmation;

                return form.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }
    }
}
