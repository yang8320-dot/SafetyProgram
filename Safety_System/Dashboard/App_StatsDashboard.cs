/// FILE: Safety_System/Dashboard/App_StatsDashboard.cs ///
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
    public class App_StatsDashboard
    {
        private string _menuDbName; 
        private Panel _mainScrollPanel;
        private TableLayoutPanel _tlpThemesContainer; 
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 資料庫表名
        private const string TblThemes = "StatsDashboard_Themes";
        private const string TblConfigs = "StatsDashboard_Configs";
        private const string TblRecords = "StatsDashboard_Records";

        // 動態區塊的 UI 封裝
        private class ThemeSectionUI
        {
            public int ThemeId { get; set; }
            public string ThemeName { get; set; }
            public GroupBox MainBox { get; set; }
            public ComboBox CboStartYear, CboStartMonth, CboEndYear, CboEndMonth;
            public DataGridView Dgv { get; set; }
        }

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName;
        }

        // 專門用於設定對話框的資料模型
        private class EditingConfig {
            public string Name { get; set; }
            public string Formula { get; set; }
        }

        public App_StatsDashboard(string menuDbName)
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
            Label lblTitle = new Label { Text = "📊 動態統計看板管理系統", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true };
            
            Button btnAddTheme = new Button { Text = "➕ 新增主題統計區塊", Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Dock = DockStyle.Right };
            btnAddTheme.FlatAppearance.BorderSize = 0;
            btnAddTheme.Click += BtnAddTheme_Click;

            pnlHeader.Controls.Add(btnAddTheme);
            pnlHeader.Controls.Add(lblTitle);
            masterLayout.Controls.Add(pnlHeader, 0, 0);

            // ================= Container =================
            _tlpThemesContainer = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                ColumnCount = 1,
                Padding = new Padding(0)
            };
            _tlpThemesContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); 

            masterLayout.Controls.Add(_tlpThemesContainer, 0, 1);
            _mainScrollPanel.Controls.Add(masterLayout);

            LoadThemes();

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
            } catch (Exception ex) { MessageBox.Show("新增失敗：" + ex.Message); }
        }

        private void LoadThemes()
        {
            _tlpThemesContainer.Controls.Clear();
            _tlpThemesContainer.RowStyles.Clear();

            DataTable dtThemes = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    // 🟢 嚴格綁定 MenuName，確保不同主選單間的主題隔離
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

            ui.MainBox = new GroupBox { 
                Text = $"📌 {themeName}", 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.DarkSlateBlue, 
                Padding = new Padding(15), 
                Margin = new Padding(0, 0, 0, 25), 
                AutoSize = true,
                Dock = DockStyle.Fill 
            };
            
            ui.MainBox.Paint += (s, e) => {
                if (ui.Dgv != null) {
                    int gridH = ui.Dgv.ColumnHeadersHeight;
                    foreach (DataGridViewRow r in ui.Dgv.Rows) gridH += r.Height;
                    ui.Dgv.Height = gridH + 5;
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ====== 第一行：操作列 ======
            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 5, 0, 15) };
            
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
            // 🟢 重新統計按鈕加寬至 165
            Button btnRecalc = new Button { Text = "🔄 重新統計", Size = new Size(165, 36), BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnRecalc.FlatAppearance.BorderSize = 0;
            Button btnSave = new Button { Text = "💾 儲存", Size = new Size(100, 36), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnSave.FlatAppearance.BorderSize = 0;
            Button btnSettings = new Button { Text = "⚙️ 顯示設定", Size = new Size(130, 36), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnSettings.FlatAppearance.BorderSize = 0;
            
            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(120, 36), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnPdf.FlatAppearance.BorderSize = 0;
            Button btnExcel = new Button { Text = "📤 導出 Excel", Size = new Size(130, 36), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 2, 0, 0) }; btnExcel.FlatAppearance.BorderSize = 0;

            // 🟢 刪除按鍵總寬度固定為 80
            Button btnDelTheme = new Button { Text = "🗑️", Size = new Size(80, 36), BackColor = Color.LightCoral, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(15, 2, 0, 0) }; btnDelTheme.FlatAppearance.BorderSize = 0;

            flpControls.Controls.AddRange(new Control[] {
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 10, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartMonth, new Label { Text = "月 ~ ", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 20, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearch, btnRecalc, btnSave, btnSettings, btnPdf, btnExcel, btnDelTheme
            });
            tlp.Controls.Add(flpControls, 0, 0);

            // ====== 第二行：資料表格 ======
            ui.Dgv = new DataGridView {
                Dock = DockStyle.Top, Height = 100, BackgroundColor = Color.White,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 12F),
                BorderStyle = BorderStyle.None, CellBorderStyle = DataGridViewCellBorderStyle.Single
            };
            ui.Dgv.EnableHeadersVisualStyles = false;
            ui.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SlateGray;
            ui.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            ui.Dgv.ColumnHeadersHeight = 40;
            ui.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            ui.Dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            ui.Dgv.Columns.Add("項目", "項目"); ui.Dgv.Columns["項目"].ReadOnly = true; ui.Dgv.Columns["項目"].FillWeight = 20;
            ui.Dgv.Columns.Add("數據", "數據資料"); ui.Dgv.Columns["數據"].ReadOnly = true; ui.Dgv.Columns["數據"].FillWeight = 40;
            ui.Dgv.Columns.Add("附件檔案", "附件檔案"); ui.Dgv.Columns["附件檔案"].ReadOnly = true; ui.Dgv.Columns["附件檔案"].FillWeight = 20;
            ui.Dgv.Columns["附件檔案"].DefaultCellStyle.ForeColor = Color.Blue; ui.Dgv.Columns["附件檔案"].DefaultCellStyle.Font = new Font(ui.Dgv.Font, FontStyle.Underline);
            ui.Dgv.Columns.Add("備註", "備註"); ui.Dgv.Columns["備註"].FillWeight = 20;

            ui.Dgv.CellFormatting += Dgv_CellFormatting;
            ui.Dgv.CellClick += (s, e) => Dgv_CellClick(s, e, ui);

            tlp.Controls.Add(ui.Dgv, 0, 1);
            ui.MainBox.Controls.Add(tlp);
            
            _tlpThemesContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _tlpThemesContainer.Controls.Add(ui.MainBox);

            // ====== 綁定事件 ======
            btnSearch.Click += async (s, e) => await CalculateAndLoadGrid(ui);
            btnRecalc.Click += async (s, e) => await RecalculateGridData(ui); 
            btnSave.Click += (s, e) => SaveGridData(ui);
            btnSettings.Click += (s, e) => { OpenSettingsDialog(ui); _ = CalculateAndLoadGrid(ui); };
            btnPdf.Click += (s, e) => ExportToPdf(ui);
            btnExcel.Click += (s, e) => ExportToExcel(ui);
            btnDelTheme.Click += (s, e) => {
                if (MessageBox.Show($"確定要刪除主題【{themeName}】及內部所有設定與資料嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            // 🟢 刪除也嚴格綁定 ThemeId，互不干擾
                            new SQLiteCommand($"DELETE FROM {TblThemes} WHERE Id={themeId}", conn).ExecuteNonQuery();
                            new SQLiteCommand($"DELETE FROM {TblConfigs} WHERE ThemeId={themeId}", conn).ExecuteNonQuery();
                            new SQLiteCommand($"DELETE FROM {TblRecords} WHERE ThemeId={themeId}", conn).ExecuteNonQuery();
                        }
                        LoadThemes();
                    } catch { }
                }
            };
        }

        // ==========================================
        // 核心公式運算引擎
        // ==========================================
        private string EvaluateStatsFormula(string template, string startYM, string endYM, Dictionary<string, DataTable> tableCache)
        {
            string evalText = template;
            Regex crossRegex = new Regex(@"(?<agg>SUM|AVG|MAX|MIN|COUNT)\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\](?:\.\[(?<dateCol>[^\]]+)\])?\)");
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

                    string cacheKey = $"{fDb}|{fTb}";
                    if (!tableCache.ContainsKey(cacheKey)) {
                        tableCache[cacheKey] = DataManager.GetTableData(fDb, fTb, "", "", "");
                    }

                    double computedVal = 0;
                    DataTable fDt = tableCache[cacheKey];
                    
                    if (fDt != null && fDt.Columns.Contains(fCol)) {
                        
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
                            if (string.IsNullOrEmpty(dateCol)) return true; 

                            string dVal = r[dateCol]?.ToString().Trim() ?? "";
                            if (string.IsNullOrEmpty(dVal)) return false;
                            
                            if (dateCol == "年度") {
                                string yearVal = dVal.Replace("年", "").Trim();
                                string sYear = startYM.Substring(0, 4);
                                string eYear = endYM.Substring(0, 4);
                                return string.Compare(yearVal, sYear) >= 0 && string.Compare(yearVal, eYear) <= 0;
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
                                return (string.Compare(compYm, startYM) >= 0 && string.Compare(compYm, endYM) <= 0);
                            }
                        });

                        List<double> vals = new List<double>();
                        foreach (var fr in matchedRows) {
                            if (double.TryParse(fr[fCol]?.ToString().Replace(",", ""), out double v)) vals.Add(v);
                        }

                        if (vals.Count > 0 || agg == "COUNT") {
                            if (agg == "SUM") computedVal = vals.Sum();
                            else if (agg == "AVG") computedVal = vals.Average();
                            else if (agg == "MAX") computedVal = vals.Max();
                            else if (agg == "MIN") computedVal = vals.Min();
                            else if (agg == "COUNT") computedVal = vals.Count;
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

        // ==========================================
        // 🟢 讀取模式 (區間交集演算法：合併附件與備註)
        // ==========================================
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
                        // 🟢 嚴格綁定 ThemeId，確保不同主題間不會互相讀取
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblConfigs} WHERE ThemeId=@T", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtConfigs);
                        }
                        
                        // 🟢 抓取該主題下的「所有歷史紀錄」，我們在記憶體中進行交集篩選
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblRecords} WHERE ThemeId=@T", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtRecords);
                        }
                    }
                } catch { return; }

                // 🟢 使用 Dictionary + HashSet 自動過濾重複的附件與備註
                Dictionary<string, (HashSet<string> Atts, HashSet<string> Rems)> mergedHistory = new Dictionary<string, (HashSet<string>, HashSet<string>)>();

                foreach (DataRow r in dtRecords.Rows) {
                    string recS = r["PeriodStart"].ToString();
                    string recE = r["PeriodEnd"].ToString();
                    string itemName = r["ItemName"].ToString();

                    // 判斷日期區間是否有交集 (Overlap: recStart <= queryEnd AND recEnd >= queryStart)
                    if (string.Compare(recS, endYM) <= 0 && string.Compare(recE, startYM) >= 0) {
                        
                        if (!mergedHistory.ContainsKey(itemName)) {
                            mergedHistory[itemName] = (new HashSet<string>(), new HashSet<string>());
                        }

                        // 合併附件
                        string att = r["Attachment"]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(att)) {
                            foreach (string p in att.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)) {
                                mergedHistory[itemName].Atts.Add(p.Trim());
                            }
                        }

                        // 合併備註
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

                    // 呼叫共用的公式運算引擎，計算最新數據
                    string evalText = EvaluateStatsFormula(template, startYM, endYM, tableCache);

                    // 🟢 寫入合併後的附件與備註
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

            ui.MainBox.Invalidate(); 
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        // ==========================================
        // 🟢 重新統計模式 (保留 Grid 上編輯中的備註與附件，僅重算數據)
        // ==========================================
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
                        
                        // 🟢 刪除舊資料時，嚴格綁定 ThemeId 避免誤刪其他主題的資料
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

        // ==========================================
        // UI 事件與附檔管理
        // ==========================================
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
                    // 🟢 附件資料夾結構加入日期區間以確保 Traceability
                    string targetFolder = $"{ui.ThemeId}/{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}_{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}"; 
                    
                    using (var frm = new AttachmentManagerUI(currentVal, "StatsDashboard", ui.ThemeId.ToString(), targetFolder, delegate(string path) { DeletePhysicalFile(path, ui.Dgv, e.RowIndex); })) {
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

        // ==========================================
        // 🟢 設定視窗：套用尺寸與間距優化，加入排序功能
        // ==========================================
        private void OpenSettingsDialog(ThemeSectionUI ui)
        {
            // 🟢 Form 加寬至 1380 容納變數產生器的新寬度
            using (Form f = new Form { Text = $"⚙️ 設定顯示與公式 - {ui.ThemeName}", Size = new Size(1380, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                
                Button btnAddNew = new Button { Text = "➕ 新增統計項目", Dock = DockStyle.Top, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
                btnAddNew.FlatAppearance.BorderSize = 0;

                Label l1 = new Label { Text = "清單項目 (拖曳可排序)", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 35, Padding = new Padding(0,10,0,0) };
                
                // 🟢 左側選單：增加雙緩衝防閃爍，並允許滑鼠拖放排序
                DataGridView dgvItems = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows=false, RowHeadersVisible=false, ColumnHeadersVisible=false, SelectionMode=DataGridViewSelectionMode.FullRowSelect, BackgroundColor=Color.White, AllowDrop=true, MultiSelect=false };
                dgvItems.Columns.Add("Name", "Name"); dgvItems.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; dgvItems.Columns["Name"].ReadOnly = true;
                
                typeof(DataGridView).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, null, dgvItems, new object[] { true });

                // ======= 下方按鈕區塊 (包含 上移、下移、刪除) =======
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

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                // 🟢 加寬上方容器
                Panel pName = new Panel { Width = 1000, Height = 45 };
                pName.Controls.Add(new Label { Text = "項目名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 350, Location = new Point(120, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);
                flpEditor.Controls.Add(pName);

                // 🟢 變數產生器 高度設定為 120 (加高) 以容納第二排
                GroupBox boxBuilder = new GroupBox { Text = "變數產生器 (自動產生跨表聚合公式)", Width = 1000, Height = 120, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding = new Padding(10) };
                Panel pnlBuilder = new Panel { Dock = DockStyle.Fill };
                
                // 第一排：庫、表、數值欄、日期欄 (🟢 調整 Y 座標與間距)
                pnlBuilder.Controls.Add(new Label { Text = "庫:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                ComboBox cbDb = new ComboBox { Location = new Point(45, 12), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                pnlBuilder.Controls.Add(cbDb);

                pnlBuilder.Controls.Add(new Label { Text = "表:", Location = new Point(195, 15), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                ComboBox cbTb = new ComboBox { Location = new Point(230, 12), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                pnlBuilder.Controls.Add(cbTb);

                pnlBuilder.Controls.Add(new Label { Text = "數值欄:", Location = new Point(440, 15), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                // 🟢 +20px 間距
                ComboBox cbCol = new ComboBox { Location = new Point(515, 12), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                pnlBuilder.Controls.Add(cbCol);

                pnlBuilder.Controls.Add(new Label { Text = "日期欄:", Location = new Point(685, 15), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                // 🟢 +20px 間距
                ComboBox cbDateCol = new ComboBox { Location = new Point(760, 12), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                pnlBuilder.Controls.Add(cbDateCol);

                // 第二排：動作、插入按鈕 (🟢 調整 Y 座標使其可見)
                pnlBuilder.Controls.Add(new Label { Text = "動作:", Location = new Point(10, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                ComboBox cbAction = new ComboBox { Location = new Point(65, 57), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                cbAction.Items.AddRange(new string[] { "加總 (SUM)", "平均值 (AVG)", "最大值 (MAX)", "最小值 (MIN)", "計數 (COUNT)" }); 
                cbAction.SelectedIndex = 0;
                pnlBuilder.Controls.Add(cbAction);

                Button btnInsert = new Button { Text = "插入公式 ⬇️", Width = 120, Height = 35, Location = new Point(220, 56), BackColor = Color.DarkCyan, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnInsert.FlatAppearance.BorderSize = 0;
                pnlBuilder.Controls.Add(btnInsert);

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
                    cbCol.Items.Clear();
                    cbDateCol.Items.Clear();
                    cbDateCol.Items.Add("[自動判斷]");

                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                        var cols = DataManager.GetColumnNames(db.EnName, tb.EnName);
                        
                        foreach(var c in cols.Where(x => x != "Id" && !x.Contains("日期") && !x.Contains("年月"))) {
                            cbCol.Items.Add(c);
                        }

                        foreach(var c in cols) cbDateCol.Items.Add(c);

                        if (cols.Contains("日期")) cbDateCol.SelectedItem = "日期";
                        else if (cols.Contains("年月")) cbDateCol.SelectedItem = "年月";
                        else if (cols.Contains("清運日期")) cbDateCol.SelectedItem = "清運日期";
                        else if (cols.Contains("年度")) cbDateCol.SelectedItem = "年度";
                        else cbDateCol.SelectedIndex = 0;
                    }
                };

                boxBuilder.Controls.Add(pnlBuilder);
                flpEditor.Controls.Add(boxBuilder);

                Label lblDesc = new Label { Text = "混合圖文公式編輯區：\n(請將純文字打在外面，將要數學計算的聚合公式包在 { 大括號 } 裡面)", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0,10,0,5), ForeColor=Color.DarkMagenta };
                
                // 🟢 同步加寬快捷鍵與輸入框 (920 -> 1000)
                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Width=1000, Height = 40, WrapContents = false };
                string[] keys = { "+", "-", "*", "/", "(", ")", "{", "}" };
                RichTextBox rtbFormula = new RichTextBox { Width=1000, Height=190, Font = new Font("Consolas", 14F), BackColor = Color.AliceBlue, Margin = new Padding(0, 5, 0, 0) };
                
                foreach (var k in keys) {
                    Button b = new Button { Text = k, Width = 45, Height = 35, Font=new Font("Consolas", 14F, FontStyle.Bold), Cursor=Cursors.Hand, BackColor=Color.WhiteSmoke };
                    pnlKeys.Controls.Add(b);
                    b.Click += (s, e) => { rtbFormula.Focus(); rtbFormula.SelectedText = $" {b.Text} "; };
                }

                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇庫、表、數值欄位！"); return; }
                    
                    string actionText = cbAction.Text;
                    string aggCode = "SUM";
                    if (actionText.Contains("SUM")) aggCode = "SUM";
                    else if (actionText.Contains("AVG")) aggCode = "AVG";
                    else if (actionText.Contains("MAX")) aggCode = "MAX";
                    else if (actionText.Contains("MIN")) aggCode = "MIN";
                    else if (actionText.Contains("COUNT")) aggCode = "COUNT";

                    string dateColSyntax = "";
                    if (cbDateCol.SelectedItem != null && cbDateCol.SelectedItem.ToString() != "[自動判斷]") {
                        dateColSyntax = $".[{(cbDateCol.SelectedItem)}]";
                    }

                    rtbFormula.Focus();
                    rtbFormula.SelectedText = $"{aggCode}([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}]{dateColSyntax})";
                };

                flpEditor.Controls.Add(lblDesc);
                flpEditor.Controls.Add(pnlKeys);
                flpEditor.Controls.Add(rtbFormula);

                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(10) };
                Button btnSaveAll = new Button { Text = "💾 儲存全部設定並關閉", Dock = DockStyle.Right, Width = 250, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor=Cursors.Hand };
                
                Button btnExp = new Button { Text = "📤 匯出設定", Dock = DockStyle.Left, Width = 130, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnImp = new Button { Text = "📥 匯入設定", Dock = DockStyle.Left, Width = 130, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                pnlBottom.Controls.Add(btnImp); pnlBottom.Controls.Add(btnExp); pnlBottom.Controls.Add(btnSaveAll);
                f.Controls.Add(tlp); f.Controls.Add(pnlBottom);

                var configs = new List<EditingConfig>();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        // 🟢 載入設定時，嚴格綁定 ThemeId
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

                // 🟢 實作排序與拖曳核心事件 (針對 dgvItems 與 configs 的同步)
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
                        // 同步調整 List 記憶體順序
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
                            // 在目標列的上方畫出紅色分割線
                            e.Graphics.DrawLine(pen, r.Left, r.Top, r.Right, r.Top);
                        }
                    }
                };

                // 🟢 實作上下按鈕事件
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

                btnSaveAll.Click += (s, e) => {
                    btnSaveAll.Focus(); 
                    
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                // 🟢 刪除舊資料時，嚴格綁定 ThemeId 避免誤刪其他主題的資料
                                new SQLiteCommand($"DELETE FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn, trans).ExecuteNonQuery();
                                
                                string sql = $"INSERT INTO {TblConfigs} (ThemeId, ItemName, FormulaTemplate) VALUES ({ui.ThemeId}, @N, @F)";
                                // 依據 Grid 目前的順序寫入資料庫，確保儲存的是調整後的順序
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

                refreshList();
                if (dgvItems.Rows.Count > 0) dgvItems.Rows[0].Selected = true;
                
                f.ShowDialog();
            }
        }

        // ==========================================
        // 匯出 PDF 與 Excel
        // ==========================================
        private void ExportToPdf(ThemeSectionUI ui)
        {
            if (ui.Dgv.Rows.Count == 0) { MessageBox.Show("目前沒有數據可供導出。"); return; }
            string dateStr = $"結算區間：{ui.CboStartYear.Text}/{ui.CboStartMonth.Text} ~ {ui.CboEndYear.Text}/{ui.CboEndMonth.Text}";
            PdfHelper.ExportDataGridViewToPdf(ui.Dgv, $"【統計看板】{ui.ThemeName}", ui.ThemeName, false, true);
        }

        private void ExportToExcel(ThemeSectionUI ui)
        {
            if (ui.Dgv.Rows.Count == 0) { MessageBox.Show("目前沒有資料可供匯出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn col in ui.Dgv.Columns) dt.Columns.Add(col.HeaderText.Replace("\n", ""));
            
            foreach (DataGridViewRow row in ui.Dgv.Rows) {
                if (row.IsNewRow) continue;
                DataRow dRow = dt.NewRow();
                for (int i = 0; i < ui.Dgv.Columns.Count; i++) {
                    dRow[i] = row.Cells[i].Value?.ToString() ?? "";
                }
                dt.Rows.Add(dRow);
            }
            ExcelHelper.ExportToExcelOrCsv(dt, ui.ThemeName, null, null, null);
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
