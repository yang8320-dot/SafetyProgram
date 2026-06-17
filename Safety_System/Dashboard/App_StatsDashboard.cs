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
        private FlowLayoutPanel _flpThemesContainer;
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
                    // 主題表
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblThemes}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, MenuName TEXT, ThemeName TEXT);", conn).ExecuteNonQuery();
                    // 設定表 (項目與公式)
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblConfigs}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, ThemeId INTEGER, ItemName TEXT, FormulaTemplate TEXT);", conn).ExecuteNonQuery();
                    // 存檔紀錄表 (綁定區間與項目)
                    new SQLiteCommand($"CREATE TABLE IF NOT EXISTS [{TblRecords}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, ThemeId INTEGER, PeriodStart TEXT, PeriodEnd TEXT, ItemName TEXT, DataValue TEXT, Attachment TEXT, Remarks TEXT);", conn).ExecuteNonQuery();
                }
            } catch { }
        }

        public Control GetView()
        {
            InitDatabase();

            _mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };

            TableLayoutPanel masterLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2, Margin = new Padding(0) };
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
            _flpThemesContainer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            _flpThemesContainer.Resize += (s, e) => {
                foreach (Control c in _flpThemesContainer.Controls) c.Width = _flpThemesContainer.ClientSize.Width - 10;
            };

            masterLayout.Controls.Add(_flpThemesContainer, 0, 1);
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
            _flpThemesContainer.Controls.Clear();

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
                _flpThemesContainer.Controls.Add(new Label { Text = "目前沒有任何統計主題，請點擊上方按鈕新增。", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Margin = new Padding(20) });
                return;
            }

            foreach (DataRow row in dtThemes.Rows) {
                BuildThemeSection(Convert.ToInt32(row["Id"]), row["ThemeName"].ToString());
            }
        }

        private void BuildThemeSection(int themeId, string themeName)
        {
            ThemeSectionUI ui = new ThemeSectionUI { ThemeId = themeId, ThemeName = themeName };

            ui.MainBox = new GroupBox { Text = $"📌 {themeName}", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Padding = new Padding(15), Margin = new Padding(0, 0, 0, 25), AutoSize = true };
            ui.MainBox.Paint += (s, e) => {
                if (ui.Dgv != null) {
                    int gridH = ui.Dgv.ColumnHeadersHeight;
                    foreach (DataGridViewRow r in ui.Dgv.Rows) gridH += r.Height;
                    ui.Dgv.Height = gridH + 5;
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ====== 第一行：操作列 ======
            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 5, 0, 15) };
            
            ui.CboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 90, Margin = new Padding(10, 4, 5, 0) };
            ui.CboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60, Margin = new Padding(0, 4, 5, 0) };
            ui.CboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 90, Margin = new Padding(0, 4, 5, 0) };
            ui.CboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60, Margin = new Padding(0, 4, 20, 0) };

            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) { ui.CboStartYear.Items.Add(i.ToString()); ui.CboEndYear.Items.Add(i.ToString()); }
            for (int i = 1; i <= 12; i++) { ui.CboStartMonth.Items.Add(i.ToString("D2")); ui.CboEndMonth.Items.Add(i.ToString("D2")); }
            
            ui.CboStartYear.SelectedItem = currY.ToString(); ui.CboStartMonth.SelectedItem = "01";
            ui.CboEndYear.SelectedItem = currY.ToString(); ui.CboEndMonth.SelectedItem = DateTime.Today.Month.ToString("D2");

            Button btnSearch = new Button { Text = "🔍 讀取 / 結算", Size = new Size(140, 36), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat }; btnSearch.FlatAppearance.BorderSize = 0;
            Button btnSave = new Button { Text = "💾 儲存表格資料", Size = new Size(160, 36), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10,0,0,0) }; btnSave.FlatAppearance.BorderSize = 0;
            Button btnSettings = new Button { Text = "⚙️ 顯示與公式設定", Size = new Size(180, 36), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10,0,0,0) }; btnSettings.FlatAppearance.BorderSize = 0;
            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(130, 36), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10,0,0,0) }; btnPdf.FlatAppearance.BorderSize = 0;
            Button btnExcel = new Button { Text = "📤 導出 Excel", Size = new Size(140, 36), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10,0,0,0) }; btnExcel.FlatAppearance.BorderSize = 0;

            Button btnDelTheme = new Button { Text = "🗑️", Size = new Size(40, 36), BackColor = Color.LightCoral, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(20,0,0,0) }; btnDelTheme.FlatAppearance.BorderSize = 0;

            flpControls.Controls.AddRange(new Control[] {
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0,8,0,0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0,8,5,0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboStartMonth, new Label { Text = "月 ~ ", AutoSize = true, Margin = new Padding(0,8,5,0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0,8,5,0), Font = new Font("Microsoft JhengHei UI", 12F) },
                ui.CboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0,8,20,0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearch, btnSave, btnSettings, btnPdf, btnExcel, btnDelTheme
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
            ui.Dgv.Columns.Add("數據", "數據 (依公式動態生成)"); ui.Dgv.Columns["數據"].ReadOnly = true; ui.Dgv.Columns["數據"].FillWeight = 40;
            ui.Dgv.Columns.Add("附件檔案", "附件檔案"); ui.Dgv.Columns["附件檔案"].ReadOnly = true; ui.Dgv.Columns["附件檔案"].FillWeight = 20;
            ui.Dgv.Columns["附件檔案"].DefaultCellStyle.ForeColor = Color.Blue; ui.Dgv.Columns["附件檔案"].DefaultCellStyle.Font = new Font(ui.Dgv.Font, FontStyle.Underline);
            ui.Dgv.Columns.Add("備註", "備註 (可自行修改)"); ui.Dgv.Columns["備註"].FillWeight = 20;

            ui.Dgv.CellFormatting += Dgv_CellFormatting;
            ui.Dgv.CellClick += (s, e) => Dgv_CellClick(s, e, ui);

            tlp.Controls.Add(ui.Dgv, 0, 1);
            ui.MainBox.Controls.Add(tlp);
            _flpThemesContainer.Controls.Add(ui.MainBox);

            // ====== 綁定事件 ======
            btnSearch.Click += async (s, e) => await CalculateAndLoadGrid(ui);
            btnSave.Click += (s, e) => SaveGridData(ui);
            btnSettings.Click += (s, e) => { OpenSettingsDialog(ui); _ = CalculateAndLoadGrid(ui); };
            btnPdf.Click += (s, e) => ExportToPdf(ui);
            btnExcel.Click += (s, e) => ExportToExcel(ui);
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

        // ==========================================
        // 核心：表格計算與存讀取邏輯
        // ==========================================
        private async Task CalculateAndLoadGrid(ThemeSectionUI ui)
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            string startYM = $"{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}";
            string endYM = $"{ui.CboEndYear.Text}-{ui.CboEndMonth.Text}";

            ui.Dgv.Rows.Clear();

            await Task.Run(() => {
                // 1. 取得該主題設定的項目與公式
                DataTable dtConfigs = new DataTable();
                DataTable dtRecords = new DataTable();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblConfigs} WHERE ThemeId=@T", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtConfigs);
                        }
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {TblRecords} WHERE ThemeId=@T AND PeriodStart=@PS AND PeriodEnd=@PE", conn)) {
                            cmd.Parameters.AddWithValue("@T", ui.ThemeId);
                            cmd.Parameters.AddWithValue("@PS", startYM);
                            cmd.Parameters.AddWithValue("@PE", endYM);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtRecords);
                        }
                    }
                } catch { return; }

                Regex crossRegex = new Regex(@"(?<agg>SUM|AVG|MAX|MIN|COUNT)\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
                Regex mathBlockRegex = new Regex(@"\{(?<expr>[^\}]+)\}");

                // 快取外部資料表避免重複 I/O
                Dictionary<string, DataTable> tableCache = new Dictionary<string, DataTable>();

                foreach (DataRow cfgRow in dtConfigs.Rows)
                {
                    string itemName = cfgRow["ItemName"].ToString();
                    string template = cfgRow["FormulaTemplate"].ToString();

                    // 找出該筆過去是否存檔過
                    DataRow savedRecord = null;
                    foreach(DataRow r in dtRecords.Rows) {
                        if (r["ItemName"].ToString() == itemName) { savedRecord = r; break; }
                    }

                    string evalText = template;

                    // 解析 { } 中的數學運算
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

                            string cacheKey = $"{fDb}|{fTb}";
                            if (!tableCache.ContainsKey(cacheKey)) {
                                tableCache[cacheKey] = DataManager.GetTableData(fDb, fTb, "", "", "");
                            }

                            double computedVal = 0;
                            DataTable fDt = tableCache[cacheKey];
                            
                            if (fDt != null && fDt.Columns.Contains(fCol)) {
                                string dateCol = fDt.Columns.Contains("年月") ? "年月" : (fDt.Columns.Contains("日期") ? "日期" : "");
                                
                                var matchedRows = fDt.Rows.Cast<DataRow>().Where(r => {
                                    if (r.RowState == DataRowState.Deleted) return false;
                                    if (string.IsNullOrEmpty(dateCol)) return true; // 若無日期欄位，全算
                                    
                                    string dVal = r[dateCol]?.ToString().Trim() ?? "";
                                    if (string.IsNullOrEmpty(dVal)) return false;
                                    
                                    // 判斷日期是否在區間內
                                    string compVal = dVal.Length >= 7 ? dVal.Substring(0,7) : dVal; // 轉為 yyyy-MM
                                    return (string.Compare(compVal, startYM) >= 0 && string.Compare(compVal, endYM) <= 0);
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
                            // 替換掉聚集函數字串為實際數字
                            expr = expr.Replace(mCross.Value, computedVal.ToString());
                        }

                        // 計算數學式
                        try {
                            DataTable dtMath = new DataTable();
                            object result = dtMath.Compute(expr, null);
                            if (result != DBNull.Value) {
                                double dRes = Convert.ToDouble(result);
                                // 預設四捨五入小數點第二位
                                evalText = evalText.Replace(mBlock.Value, Math.Round(dRes, 2).ToString("0.##"));
                            }
                        } catch {
                            evalText = evalText.Replace(mBlock.Value, "NaN");
                        }
                    }

                    string attach = savedRecord != null ? savedRecord["Attachment"].ToString() : "";
                    string remarks = savedRecord != null ? savedRecord["Remarks"].ToString() : "";

                    if (ui.Dgv.InvokeRequired) {
                        ui.Dgv.Invoke(new Action(() => ui.Dgv.Rows.Add(itemName, evalText, attach, remarks)));
                    } else {
                        ui.Dgv.Rows.Add(itemName, evalText, attach, remarks);
                    }
                }
            });

            ui.MainBox.Invalidate(); 
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
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
                    string targetFolder = $"{ui.CboStartYear.Text}-{ui.CboStartMonth.Text}"; 
                    
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
        // 設定視窗 (自訂項目與公式)
        // ==========================================
        private void OpenSettingsDialog(ThemeSectionUI ui)
        {
            using (Form f = new Form { Text = $"⚙️ 設定顯示與公式 - {ui.ThemeName}", Size = new Size(1300, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                
                Button btnAddNew = new Button { Text = "➕ 新增統計項目", Dock = DockStyle.Top, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
                btnAddNew.FlatAppearance.BorderSize = 0;

                Label l1 = new Label { Text = "清單項目 (拖曳可排序)", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 35, Padding = new Padding(0,10,0,0) };
                
                // 改用 DataGridView 方便拖曳排序
                DataGridView dgvItems = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows=false, RowHeadersVisible=false, ColumnHeadersVisible=false, SelectionMode=DataGridViewSelectionMode.FullRowSelect, BackgroundColor=Color.White, AllowDrop=true, MultiSelect=false };
                dgvItems.Columns.Add("Name", "Name"); dgvItems.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; dgvItems.Columns["Name"].ReadOnly = true;
                
                int dragFromIdx = -1; Rectangle dragBox = Rectangle.Empty;
                dgvItems.MouseDown += (s, e) => { var hit = dgvItems.HitTest(e.X, e.Y); if (hit.RowIndex >= 0) { dragFromIdx = hit.RowIndex; dragBox = new Rectangle(new Point(e.X - 10, e.Y - 10), new Size(20,20)); } else dragBox = Rectangle.Empty; };
                dgvItems.MouseMove += (s, e) => { if ((e.Button & MouseButtons.Left) == MouseButtons.Left && dragBox != Rectangle.Empty && !dragBox.Contains(e.X, e.Y)) dgvItems.DoDragDrop(dgvItems.Rows[dragFromIdx], DragDropEffects.Move); };
                dgvItems.DragOver += (s, e) => e.Effect = DragDropEffects.Move;
                dgvItems.DragDrop += (s, e) => {
                    Point p = dgvItems.PointToClient(new Point(e.X, e.Y)); var hit = dgvItems.HitTest(p.X, p.Y);
                    int targetIdx = hit.RowIndex; if (targetIdx < 0) targetIdx = dgvItems.Rows.Count - 1;
                    if (dragFromIdx >= 0 && dragFromIdx != targetIdx) {
                        var row = dgvItems.Rows[dragFromIdx]; dgvItems.Rows.RemoveAt(dragFromIdx); dgvItems.Rows.Insert(targetIdx, row); dgvItems.ClearSelection(); dgvItems.Rows[targetIdx].Selected = true;
                    }
                };

                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                pnlLeft.Controls.Add(dgvItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnAddNew);
                pnlLeft.Controls.Add(btnDel);

                // ============== 右側編輯區 ==============
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯選取項目的內容公式", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                Panel pName = new Panel { Width = 950, Height = 45 };
                pName.Controls.Add(new Label { Text = "項目名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 350, Location = new Point(100, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);
                flpEditor.Controls.Add(pName);

                // 公式變數生成器
                GroupBox boxBuilder = new GroupBox { Text = "變數產生器 (自動產生跨表聚合公式)", Width=920, Height = 100, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding=new Padding(10) };
                Panel pnlBuilder = new Panel { Dock = DockStyle.Fill };
                
                ComboBox cbAction = new ComboBox { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(10, 20) };
                cbAction.Items.AddRange(new string[] { "SUM", "AVG", "MAX", "MIN", "COUNT" }); cbAction.SelectedIndex = 0;
                pnlBuilder.Controls.Add(cbAction);
                
                ComboBox cbDb = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(135, 20) };
                ComboBox cbTb = new ComboBox { Width = 230, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(295, 20) };
                ComboBox cbCol = new ComboBox { Width = 210, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(535, 20) };
                pnlBuilder.Controls.AddRange(new Control[] { cbDb, cbTb, cbCol });

                Button btnInsert = new Button { Text = "插入 ⬇️", Width = 100, Height = 35, BackColor = Color.DarkCyan, ForeColor=Color.White, Cursor=Cursors.Hand, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), FlatStyle=FlatStyle.Flat, Location = new Point(760, 18) };
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
                    cbCol.Items.Clear(); cbCol.Items.Add("Id (無條件計數)");
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                        var cols = DataManager.GetColumnNames(db.EnName, tb.EnName);
                        foreach(var c in cols.Where(x => x != "Id" && !x.Contains("日期") && !x.Contains("年月"))) cbCol.Items.Add(c);
                    }
                };

                boxBuilder.Controls.Add(pnlBuilder);
                flpEditor.Controls.Add(boxBuilder);

                Label lblDesc = new Label { Text = "混合圖文公式編輯區：\n(請將純文字打在外面，將要數學計算的聚合公式包在 { 大括號 } 裡面)", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0,10,0,5), ForeColor=Color.DarkMagenta };
                RichTextBox rtbFormula = new RichTextBox { Width=920, Height=250, Font = new Font("Consolas", 14F), BackColor = Color.AliceBlue };
                
                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇庫、表、欄位！"); return; }
                    string cName = cbCol.SelectedItem.ToString() == "Id (無條件計數)" ? "Id" : cbCol.SelectedItem.ToString();
                    rtbFormula.Focus();
                    rtbFormula.SelectedText = $"{cbAction.Text}([{db.EnName}].[{tb.EnName}].[{cName}])";
                };

                flpEditor.Controls.Add(lblDesc);
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

                // 載入與綁定資料
                var configs = new List<(string Name, string Formula)>();
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT ItemName, FormulaTemplate FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn))
                        using (var reader = cmd.ExecuteReader()) {
                            while(reader.Read()) configs.Add((reader[0].ToString(), reader[1].ToString()));
                        }
                    }
                } catch { }

                Action refreshList = () => { dgvItems.Rows.Clear(); foreach(var c in configs) dgvItems.Rows.Add(c.Name); };
                refreshList();

                bool isSyncing = false;
                dgvItems.SelectionChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    txtName.Text = configs[idx].Name;
                    rtbFormula.Text = configs[idx].Formula;
                };

                txtName.TextChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    configs[idx] = (txtName.Text, configs[idx].Formula);
                    isSyncing = true; dgvItems.Rows[idx].Cells[0].Value = txtName.Text; isSyncing = false;
                };

                rtbFormula.TextChanged += (s, e) => {
                    if (isSyncing || dgvItems.SelectedRows.Count == 0) return;
                    int idx = dgvItems.SelectedRows[0].Index;
                    configs[idx] = (configs[idx].Name, rtbFormula.Text);
                };

                btnAddNew.Click += (s, e) => {
                    configs.Add(("新項目", "")); refreshList();
                    dgvItems.ClearSelection(); dgvItems.Rows[dgvItems.Rows.Count-1].Selected = true;
                    txtName.Focus();
                };

                btnDel.Click += (s, e) => {
                    if (dgvItems.SelectedRows.Count > 0) {
                        configs.RemoveAt(dgvItems.SelectedRows[0].Index); refreshList();
                        txtName.Clear(); rtbFormula.Clear();
                    }
                };

                btnSaveAll.Click += (s, e) => {
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                new SQLiteCommand($"DELETE FROM {TblConfigs} WHERE ThemeId={ui.ThemeId}", conn, trans).ExecuteNonQuery();
                                string sql = $"INSERT INTO {TblConfigs} (ThemeId, ItemName, FormulaTemplate) VALUES ({ui.ThemeId}, @N, @F)";
                                foreach(var c in configs) {
                                    if(string.IsNullOrWhiteSpace(c.Name)) continue;
                                    using(var cmd = new SQLiteCommand(sql, conn, trans)) {
                                        cmd.Parameters.AddWithValue("@N", c.Name); cmd.Parameters.AddWithValue("@F", c.Formula);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                trans.Commit();
                            }
                        }
                        f.DialogResult = DialogResult.OK;
                    } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
                };

                // 匯出匯入邏輯
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
                                        if(!string.IsNullOrEmpty(n)) configs.Add((n, fma));
                                    }
                                }
                                refreshList();
                                MessageBox.Show("匯入成功！請點擊右下角「儲存全部設定」以生效。");
                            } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message); }
                        }
                    }
                };

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
