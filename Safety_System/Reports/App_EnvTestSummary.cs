/// FILE: Safety_System/Reports/App_EnvTestSummary.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_EnvTestSummary
    {
        private ComboBox _cboStartYear;
        private ComboBox _cboEndYear;
        private Button _btnSearch;
        private Button _btnPdf;
        private Button _btnExcel;
        private Button _btnSettings;
        
        // 用來容納多個 DataGridView (每 5 年一個) 的容器
        private FlowLayoutPanel _flpGridsContainer;

        private const string ConfigDbName = "SystemConfig";
        private const string ConfigTableName = "EnvTestSummaryConfigs";

        // 設定檔模型
        private class EnvConfigItem
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string SegCol { get; set; }
            public string ItemCol { get; set; }
            public string LimitCol { get; set; }
            public string DateCol { get; set; }
            public string ValueCol { get; set; }
            public string NoteCol { get; set; }
        }

        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        // 用來暫存從資料庫撈出的單筆資料
        private class TestRecord
        {
            public string SEG { get; set; }
            public string Item { get; set; }
            public string Limit { get; set; }
            public int RocYear { get; set; }
            public DateTime Date { get; set; }
            public string Value { get; set; }
            public string Note { get; set; }
        }

        private List<EnvConfigItem> _configs = new List<EnvConfigItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS [EnvTestSummaryConfigs] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        DbName TEXT, TableName TEXT, SegCol TEXT, ItemCol TEXT, 
                        LimitCol TEXT, DateCol TEXT, ValueCol TEXT, NoteCol TEXT);";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public Control GetView()
        {
            InitDatabase();
            _dbMap = App_DbConfig.GetDbMapCache();
            LoadSettings();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            // ==========================================
            // 第一個框：資料選擇與操作列
            // ==========================================
            GroupBox box1 = new GroupBox { Text = "⚙️ 查詢條件與操作區", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0,0,0,20) };
            
            FlowLayoutPanel flpRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,5,0,10), WrapContents = false };
            
            _cboStartYear = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(0, 3, 5, 0) };
            _cboEndYear = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(5, 3, 0, 0) };
            
            _cboStartYear.Items.Add("全部年度");
            _cboEndYear.Items.Add("全部年度");

            int currentYear = DateTime.Today.Year;
            for (int i = currentYear - 10; i <= currentYear + 2; i++) {
                _cboStartYear.Items.Add(i.ToString());
                _cboEndYear.Items.Add(i.ToString());
            }
            
            _cboStartYear.SelectedIndex = 0; 
            _cboEndYear.SelectedIndex = 0;   

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(120, 38), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(20, 0, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += BtnSearch_Click;

            _btnSettings = new Button { Text = "⚙️ 來源設定", Size = new Size(130, 38), BackColor = Color.DimGray, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnSettings.FlatAppearance.BorderSize = 0;
            _btnSettings.Click += BtnSettings_Click;

            _btnExcel = new Button { Text = "📤 匯出 Excel", Size = new Size(140, 38), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnExcel.FlatAppearance.BorderSize = 0;
            _btnExcel.Click += BtnExcel_Click;

            _btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(140, 38), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnPdf.FlatAppearance.BorderSize = 0;
            _btnPdf.Click += BtnPdf_Click;

            flpRow.Controls.AddRange(new Control[] {
                new Label { Text = "查詢年度區間:", AutoSize = true, Margin = new Padding(10, 10, 5, 0) }, 
                _cboStartYear, 
                new Label { Text = "～", AutoSize = true, Margin = new Padding(5, 10, 5, 0) }, 
                _cboEndYear,
                _btnSearch, _btnSettings, _btnExcel, _btnPdf
            });

            box1.Controls.Add(flpRow);
            layout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 第二個框：預覽區 (動態生成多個 DGV)
            // ==========================================
            GroupBox box2 = new GroupBox { Text = "📄 環測數據一覽表 (每 5 年一個區塊)", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            _flpGridsContainer = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoScroll = true, 
                FlowDirection = FlowDirection.TopDown, 
                WrapContents = false,
                BackColor = Color.White
            };
            _flpGridsContainer.Resize += (s, e) => {
                foreach (Control c in _flpGridsContainer.Controls) {
                    c.Width = _flpGridsContainer.ClientSize.Width - 10;
                }
            };

            box2.Controls.Add(_flpGridsContainer);
            layout.Controls.Add(box2, 0, 1);

            mainScrollPanel.Controls.Add(layout);
            
            if (_configs.Count > 0) BtnSearch_Click(null, null);

            return mainScrollPanel;
        }

        private void LoadSettings()
        {
            _configs.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"SELECT * FROM {ConfigTableName}", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            _configs.Add(new EnvConfigItem { 
                                DbName = reader["DbName"].ToString(), 
                                TableName = reader["TableName"].ToString(), 
                                SegCol = reader["SegCol"].ToString(),
                                ItemCol = reader["ItemCol"].ToString(), 
                                LimitCol = reader["LimitCol"].ToString(),
                                DateCol = reader["DateCol"].ToString(), 
                                ValueCol = reader["ValueCol"].ToString(), 
                                NoteCol = reader["NoteCol"].ToString()
                            });
                        }
                    }
                }
            } catch { }
        }

        private void SaveSettings()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        new SQLiteCommand($"DELETE FROM {ConfigTableName}", conn, trans).ExecuteNonQuery();

                        string insertSql = $"INSERT INTO {ConfigTableName} (DbName, TableName, SegCol, ItemCol, LimitCol, DateCol, ValueCol, NoteCol) VALUES (@DB, @TB, @SC, @IC, @LC, @DC, @VC, @NC)";
                        foreach (var c in _configs) {
                            using (var cmd = new SQLiteCommand(insertSql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@DB", c.DbName);
                                cmd.Parameters.AddWithValue("@TB", c.TableName);
                                cmd.Parameters.AddWithValue("@SC", c.SegCol);
                                cmd.Parameters.AddWithValue("@IC", c.ItemCol);
                                cmd.Parameters.AddWithValue("@LC", c.LimitCol);
                                cmd.Parameters.AddWithValue("@DC", c.DateCol);
                                cmd.Parameters.AddWithValue("@VC", c.ValueCol);
                                cmd.Parameters.AddWithValue("@NC", c.NoteCol);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }

        // ==========================================
        // 核心資料處理與繪製邏輯
        // ==========================================
        private async void BtnSearch_Click(object sender, EventArgs e)
        {
            if (_configs.Count == 0) {
                MessageBox.Show("目前尚未設定任何資料來源，請點擊【來源設定】新增來源欄位！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            string sYearStr = _cboStartYear.SelectedItem.ToString();
            string eYearStr = _cboEndYear.SelectedItem.ToString();
            int filterStartYear = sYearStr == "全部年度" ? 0 : int.Parse(sYearStr);
            int filterEndYear = eYearStr == "全部年度" ? 9999 : int.Parse(eYearStr);

            if (filterStartYear > filterEndYear) {
                MessageBox.Show("起始年度不能大於結束年度！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                return;
            }

            _flpGridsContainer.SuspendLayout();
            _flpGridsContainer.Controls.Clear();

            try 
            {
                // 1. 撈取所有設定表的資料
                List<TestRecord> allRecords = new List<TestRecord>();

                await Task.Run(() => {
                    foreach (var cfg in _configs) {
                        DataTable dt = null;
                        try {
                            dt = DataManager.GetTableData(cfg.DbName, cfg.TableName, "", "", "");
                        } catch { continue; }

                        if (dt == null || dt.Rows.Count == 0) continue;

                        string actualSegCol = !string.IsNullOrEmpty(cfg.SegCol) && dt.Columns.Contains(cfg.SegCol) ? cfg.SegCol : "";
                        string actualItemCol = !string.IsNullOrEmpty(cfg.ItemCol) && dt.Columns.Contains(cfg.ItemCol) ? cfg.ItemCol : "";
                        string actualLimitCol = !string.IsNullOrEmpty(cfg.LimitCol) && dt.Columns.Contains(cfg.LimitCol) ? cfg.LimitCol : "";
                        string actualDateCol = !string.IsNullOrEmpty(cfg.DateCol) && dt.Columns.Contains(cfg.DateCol) ? cfg.DateCol : "";
                        string actualValueCol = !string.IsNullOrEmpty(cfg.ValueCol) && dt.Columns.Contains(cfg.ValueCol) ? cfg.ValueCol : "";
                        string actualNoteCol = !string.IsNullOrEmpty(cfg.NoteCol) && dt.Columns.Contains(cfg.NoteCol) ? cfg.NoteCol : "";

                        if (string.IsNullOrEmpty(actualItemCol) || string.IsNullOrEmpty(actualDateCol) || string.IsNullOrEmpty(actualValueCol)) continue;

                        foreach (DataRow r in dt.Rows) {
                            if (r.RowState == DataRowState.Deleted) continue;

                            string rawDate = r[actualDateCol]?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(rawDate)) continue;

                            var (parsedRocYear, parsedDate) = ParseToRocYearAndDate(rawDate);
                            if (parsedRocYear == 0) continue;

                            // 轉換為西元年進行過濾
                            int westernYear = parsedRocYear + 1911;
                            if (filterStartYear != 0 && westernYear < filterStartYear) continue;
                            if (filterEndYear != 9999 && westernYear > filterEndYear) continue;

                            string valStr = r[actualValueCol]?.ToString()?.Trim() ?? "";
                            if (string.IsNullOrEmpty(valStr)) continue;

                            allRecords.Add(new TestRecord {
                                SEG = string.IsNullOrEmpty(actualSegCol) ? "" : r[actualSegCol]?.ToString()?.Trim() ?? "",
                                Item = r[actualItemCol]?.ToString()?.Trim() ?? "",
                                Limit = string.IsNullOrEmpty(actualLimitCol) ? "" : r[actualLimitCol]?.ToString()?.Trim() ?? "",
                                RocYear = parsedRocYear,
                                Date = parsedDate,
                                Value = valStr,
                                Note = string.IsNullOrEmpty(actualNoteCol) ? "" : r[actualNoteCol]?.ToString()?.Trim() ?? ""
                            });
                        }
                    }
                });

                if (allRecords.Count == 0) {
                    Label lblEmpty = new Label { Text = "該區間內查無資料。", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(20) };
                    _flpGridsContainer.Controls.Add(lblEmpty);
                    return;
                }

                // 2. 切割 5 年區塊
                var distinctRocYears = allRecords.Select(r => r.RocYear).Distinct().OrderBy(y => y).ToList();
                int minRocYear = distinctRocYears.First();
                int maxRocYear = distinctRocYears.Last();

                List<(int StartYear, int EndYear)> chunks = new List<(int, int)>();
                int currentChunkStart = minRocYear;
                while (currentChunkStart <= maxRocYear) {
                    chunks.Add((currentChunkStart, currentChunkStart + 4));
                    currentChunkStart += 5;
                }

                // 3. 計算每年需要的欄位數 (該年度中，同一個 SEG + Item 最多有幾個不同的測項日期)
                Dictionary<int, int> maxPeriodsPerYear = new Dictionary<int, int>();
                var groupedByYear = allRecords.GroupBy(r => r.RocYear);
                
                foreach (var yg in groupedByYear) {
                    int year = yg.Key;
                    int maxPeriods = 1;
                    
                    var groupedBySegItem = yg.GroupBy(r => new { r.SEG, r.Item });
                    foreach (var sig in groupedBySegItem) {
                        // 針對同一個測項，依日期排序
                        var sortedDates = sig.Select(r => r.Date).Distinct().OrderBy(d => d).ToList();
                        if (sortedDates.Count > maxPeriods) {
                            maxPeriods = sortedDates.Count;
                        }
                    }
                    maxPeriodsPerYear[year] = maxPeriods;
                }

                // 4. 為每個 5 年區塊建立 DataTable 與 DataGridView
                foreach (var chunk in chunks) {
                    var recordsInChunk = allRecords.Where(r => r.RocYear >= chunk.StartYear && r.RocYear <= chunk.EndYear).ToList();
                    if (recordsInChunk.Count == 0) continue;

                    DataTable dtChunk = new DataTable();
                    dtChunk.Columns.Add("SEG編號", typeof(string));
                    dtChunk.Columns.Add("檢測項目", typeof(string));
                    dtChunk.Columns.Add("容許濃度標準", typeof(string));

                    // 動態產生這 5 年的欄位 (例如: 107-1, 107-2, 108-1...)
                    List<string> periodCols = new List<string>();
                    for (int y = chunk.StartYear; y <= chunk.EndYear; y++) {
                        if (maxPeriodsPerYear.ContainsKey(y)) {
                            for (int p = 1; p <= maxPeriodsPerYear[y]; p++) {
                                string colName = $"{y}-{p}";
                                dtChunk.Columns.Add(colName, typeof(string));
                                periodCols.Add(colName);
                            }
                        }
                    }
                    dtChunk.Columns.Add("備註", typeof(string));

                    // 將資料填入
                    var groupedRecords = recordsInChunk.GroupBy(r => new { r.SEG, r.Item, r.Limit })
                                                       .OrderBy(g => g.Key.SEG).ThenBy(g => g.Key.Item);

                    foreach (var group in groupedRecords) {
                        DataRow row = dtChunk.NewRow();
                        row["SEG編號"] = group.Key.SEG;
                        row["檢測項目"] = group.Key.Item;
                        row["容許濃度標準"] = group.Key.Limit;

                        // 針對每一年的紀錄進行填寫
                        var recordsByYear = group.GroupBy(r => r.RocYear);
                        foreach (var yg in recordsByYear) {
                            int year = yg.Key;
                            // 依日期排序，決定是 -1, -2, -3...
                            var sortedByDate = yg.OrderBy(r => r.Date).ToList();
                            
                            // 因為同一天可能有多筆資料（或許是更新或重測），我們依序把它接起來
                            Dictionary<DateTime, List<string>> valuesByDate = new Dictionary<DateTime, List<string>>();
                            foreach (var r in sortedByDate) {
                                if (!valuesByDate.ContainsKey(r.Date)) valuesByDate[r.Date] = new List<string>();
                                valuesByDate[r.Date].Add(r.Value);
                            }

                            int periodIndex = 1;
                            var distinctSortedDates = valuesByDate.Keys.OrderBy(d => d).ToList();
                            foreach (var d in distinctSortedDates) {
                                string colName = $"{year}-{periodIndex}";
                                if (dtChunk.Columns.Contains(colName)) {
                                    row[colName] = string.Join("\n", valuesByDate[d]);
                                }
                                periodIndex++;
                            }
                        }

                        // 彙整備註
                        var notes = group.Select(r => r.Note).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct();
                        row["備註"] = string.Join("\n", notes);

                        dtChunk.Rows.Add(row);
                    }

                    // 建立 UI 呈現
                    Panel pnlWrapper = new Panel { 
                        Width = _flpGridsContainer.ClientSize.Width - 10, 
                        BackColor = Color.White, 
                        Margin = new Padding(0, 0, 0, 30) 
                    };
                    pnlWrapper.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pnlWrapper.ClientRectangle, Color.DarkGray, ButtonBorderStyle.Solid);

                    Label lblTitle = new Label { 
                        Text = $"📊 {chunk.StartYear + 1911} ~ {chunk.EndYear + 1911} 年 (民國 {chunk.StartYear} ~ {chunk.EndYear} 年) 檢測數據", 
                        Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                        ForeColor = Color.DarkSlateBlue, 
                        AutoSize = true, 
                        Padding = new Padding(15, 10, 0, 10), 
                        Dock = DockStyle.Top 
                    };

                    DataGridView dgv = new DataGridView {
                        Dock = DockStyle.Fill,
                        BackgroundColor = Color.White,
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        ReadOnly = true,
                        RowHeadersVisible = false,
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                        Font = new Font("Microsoft JhengHei UI", 11F),
                        BorderStyle = BorderStyle.None,
                        CellBorderStyle = DataGridViewCellBorderStyle.Single
                    };
                    
                    dgv.EnableHeadersVisualStyles = false;
                    dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                    dgv.ColumnHeadersHeight = 45;
                    dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                    
                    dgv.DataSource = dtChunk;
                    
                    // 調整寬度
                    if (dgv.Columns.Contains("SEG編號")) dgv.Columns["SEG編號"].FillWeight = 10;
                    if (dgv.Columns.Contains("檢測項目")) dgv.Columns["檢測項目"].FillWeight = 15;
                    if (dgv.Columns.Contains("容許濃度標準")) dgv.Columns["容許濃度標準"].FillWeight = 10;
                    if (dgv.Columns.Contains("備註")) dgv.Columns["備註"].FillWeight = 15;
                    foreach (var c in periodCols) {
                        if (dgv.Columns.Contains(c)) dgv.Columns[c].FillWeight = 8;
                    }

                    dgv.ClearSelection();

                    // 自動計算高度
                    int totalHeight = dgv.ColumnHeadersHeight;
                    foreach (DataGridViewRow r in dgv.Rows) totalHeight += r.Height;
                    // 加上標題高度與 Padding
                    pnlWrapper.Height = totalHeight + lblTitle.Height + 15;
                    if (pnlWrapper.Height < 150) pnlWrapper.Height = 150;

                    pnlWrapper.Controls.Add(dgv);
                    pnlWrapper.Controls.Add(lblTitle);
                    
                    _flpGridsContainer.Controls.Add(pnlWrapper);
                }
            } 
            catch (Exception ex) 
            {
                MessageBox.Show("查詢失敗：" + ex.Message, "錯誤");
            } 
            finally 
            {
                _flpGridsContainer.ResumeLayout(true);
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private (int RocYear, DateTime Date) ParseToRocYearAndDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return (0, DateTime.MinValue);
            dateStr = dateStr.Trim().Replace("/", "-");

            Regex twRegex = new Regex(@"^(?<year>\d{2,3})(?:-(?<month>\d{1,2}))(?:-(?<day>\d{1,2}))(?:\s+.*)?$");
            Match matchTw = twRegex.Match(dateStr);

            if (matchTw.Success)
            {
                if (int.TryParse(matchTw.Groups["year"].Value, out int twYear) &&
                    int.TryParse(matchTw.Groups["month"].Value, out int month) &&
                    int.TryParse(matchTw.Groups["day"].Value, out int day))
                {
                    int finalYear = twYear < 200 ? twYear + 1911 : twYear;
                    try {
                        DateTime dt = new DateTime(finalYear, month, day);
                        return (finalYear - 1911, dt);
                    } catch { return (0, DateTime.MinValue); }
                }
            }

            if (DateTime.TryParse(dateStr, out DateTime result)) {
                return (result.Year - 1911, result);
            }

            return (0, DateTime.MinValue);
        }

        // ==========================================
        // 匯出 Excel
        // ==========================================
        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (_flpGridsContainer.Controls.Count == 0) {
                MessageBox.Show("畫面上沒有資料可供匯出，請先執行查詢！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = $"環測數據一覽表_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        using (ExcelPackage p = new ExcelPackage())
                        {
                            foreach (Control ctrl in _flpGridsContainer.Controls)
                            {
                                if (ctrl is Panel pnl)
                                {
                                    // 從標題抓取區間名稱
                                    string sheetName = "Data";
                                    foreach (Control c in pnl.Controls) {
                                        if (c is Label lbl && lbl.Text.Contains("年")) {
                                            // 萃取 "107 ~ 111 年" 這樣的字眼
                                            Match m = Regex.Match(lbl.Text, @"民國\s*(\d+\s*~\s*\d+)\s*年");
                                            if (m.Success) {
                                                sheetName = m.Groups[1].Value.Replace(" ", "") + "年";
                                            }
                                            break;
                                        }
                                    }

                                    DataGridView dgv = pnl.Controls.OfType<DataGridView>().FirstOrDefault();
                                    if (dgv != null && dgv.DataSource is DataTable dt) {
                                        var ws = p.Workbook.Worksheets.Add(sheetName);
                                        ws.Cells["A1"].LoadFromDataTable(dt, true);
                                        
                                        // 簡單的排版
                                        using (var range = ws.Cells[1, 1, 1, dt.Columns.Count]) {
                                            range.Style.Font.Bold = true;
                                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        }

                                        var dataRange = ws.Cells[2, 1, dt.Rows.Count + 1, dt.Columns.Count];
                                        dataRange.Style.WrapText = true;
                                        dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                        ws.Cells.AutoFitColumns();
                                    }
                                }
                            }
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("Excel 匯出成功！每個區間已獨立為一個活頁簿。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // ==========================================
        // 匯出 PDF (使用 PdfHelper 模式)
        // ==========================================
        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (_flpGridsContainer.Controls.Count == 0) {
                MessageBox.Show("畫面上沒有資料可供匯出，請先執行查詢！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try 
            {
                Application.DoEvents(); 

                List<Bitmap> bitmaps = new List<Bitmap>();
                foreach (Control ctrl in _flpGridsContainer.Controls) 
                {
                    if (ctrl is Panel pnl)
                    {
                        // 確保畫面有被完全展開繪製
                        int origHeight = pnl.Height;
                        DataGridView dgv = pnl.Controls.OfType<DataGridView>().FirstOrDefault();
                        if (dgv != null) {
                            int exactGridHeight = dgv.ColumnHeadersHeight;
                            foreach(DataGridViewRow r in dgv.Rows) exactGridHeight += r.Height;
                            pnl.Height = exactGridHeight + 60; // 加上標題空間
                        }

                        Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                        pnl.DrawToBitmap(bmp, new Rectangle(0, 0, pnl.Width, pnl.Height));
                        bitmaps.Add(bmp);

                        pnl.Height = origHeight;
                    }
                }

                string sYear = _cboStartYear.SelectedItem.ToString();
                string eYear = _cboEndYear.SelectedItem.ToString();
                string dateStr = $"查詢區間：{(sYear == "全部年度" ? "所有紀錄" : $"{sYear} ~ {eYear} 年度")}";
                
                PdfHelper.ExportDashboardToPdf(bitmaps, "環測數據一覽表", dateStr, "環測數據一覽表");
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

        // ==========================================
        // 設定畫面
        // ==========================================
        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "⚙️ 讀取資料來源設定 (定義要合併統計的資料表)", Size = new Size(1300, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Panel pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.WhiteSmoke };

                FlowLayoutPanel flpMain = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10) };

                var editingConfigs = new List<EnvConfigItem>();
                foreach (var c in _configs) editingConfigs.Add(new EnvConfigItem { 
                    DbName = c.DbName, TableName = c.TableName, SegCol = c.SegCol, ItemCol = c.ItemCol, 
                    LimitCol = c.LimitCol, DateCol = c.DateCol, ValueCol = c.ValueCol, NoteCol = c.NoteCol 
                });

                Dictionary<string, List<string>> _columnCache = new Dictionary<string, List<string>>();

                Action renderRows = null;
                renderRows = () => {
                    flpMain.SuspendLayout();
                    flpMain.Controls.Clear();

                    Panel pnlHeader = new Panel { Width = 1250, Height = 30 };
                    string[] headers = { "刪除", "資料庫", "資料表", "對應[SEG編號]", "對應[檢測項目]", "對應[管制標準]", "對應[日期]", "對應[數據值]", "對應[備註]" };
                    
                    int[] xs = { 0, 45, 175, 335, 465, 595, 725, 855, 985 };
                    int[] ws = { 35, 120, 150, 120, 120, 120, 120, 120, 120 };

                    for (int i = 0; i < 9; i++) {
                        Label lbl = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Location = new Point(xs[i], 5), AutoSize = true };
                        pnlHeader.Controls.Add(lbl);
                    }
                    flpMain.Controls.Add(pnlHeader);

                    for (int i = 0; i < editingConfigs.Count; i++) {
                        int currentIndex = i;
                        var conf = editingConfigs[i];

                        Panel pnlRow = new Panel { Width = 1250, Height = 40, Margin = new Padding(0, 2, 0, 2) };

                        Button btnDel = new Button { Text = "❌", Location = new Point(xs[0], 2), Size = new Size(ws[0], 32), BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        btnDel.FlatAppearance.BorderSize = 0;
                        btnDel.Click += (s, ev) => { editingConfigs.RemoveAt(currentIndex); renderRows(); };

                        ComboBox cbDb = new ComboBox { Location = new Point(xs[1], 5), Width = ws[1], DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbTb = new ComboBox { Location = new Point(xs[2], 5), Width = ws[2], DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbSeg = new ComboBox { Location = new Point(xs[3], 5), Width = ws[3], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbItem = new ComboBox { Location = new Point(xs[4], 5), Width = ws[4], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbLimit = new ComboBox { Location = new Point(xs[5], 5), Width = ws[5], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbDate = new ComboBox { Location = new Point(xs[6], 5), Width = ws[6], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbVal = new ComboBox { Location = new Point(xs[7], 5), Width = ws[7], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbNote = new ComboBox { Location = new Point(xs[8], 5), Width = ws[8], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };

                        cbDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                        foreach (var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

                        bool colsLoaded = false;
                        bool isInitializing = true; 

                        Action<string, string> lazyLoadCols = (dbEnName, tbEnName) => {
                            if (colsLoaded) return;
                            
                            List<string> cols = new List<string>();
                            if (!string.IsNullOrEmpty(tbEnName) && !string.IsNullOrEmpty(dbEnName)) {
                                string cacheKey = $"{dbEnName}_{tbEnName}";
                                if (_columnCache.ContainsKey(cacheKey)) {
                                    cols = _columnCache[cacheKey]; 
                                } else {
                                    cols = DataManager.GetColumnNames(dbEnName, tbEnName);
                                    _columnCache[cacheKey] = cols; 
                                }
                            }

                            string sSeg = cbSeg.Text, sItem = cbItem.Text, sLimit = cbLimit.Text, sDate = cbDate.Text, sVal = cbVal.Text, sNote = cbNote.Text;

                            ComboBox[] boxes = { cbSeg, cbItem, cbLimit, cbDate, cbVal, cbNote };
                            foreach (var b in boxes) { b.Items.Clear(); b.Items.Add(""); }

                            foreach(var c in cols) {
                                if (c != "Id") {
                                    foreach (var b in boxes) b.Items.Add(c);
                                }
                            }
                            
                            cbSeg.Text = sSeg; cbItem.Text = sItem; cbLimit.Text = sLimit; cbDate.Text = sDate; cbVal.Text = sVal; cbNote.Text = sNote;
                            colsLoaded = true;
                        };

                        EventHandler triggerLoad = (s, ev) => {
                            if (cbDb.SelectedItem != null && cbTb.SelectedItem != null) {
                                lazyLoadCols(((ItemMap)cbDb.SelectedItem).EnName, ((ItemMap)cbTb.SelectedItem).EnName);
                            }
                        };

                        cbSeg.DropDown += triggerLoad; cbItem.DropDown += triggerLoad; cbLimit.DropDown += triggerLoad;
                        cbDate.DropDown += triggerLoad; cbVal.DropDown += triggerLoad; cbNote.DropDown += triggerLoad;

                        cbDb.SelectedIndexChanged += (s, ev) => {
                            if (isInitializing) return;
                            var selDb = cbDb.SelectedItem as ItemMap;
                            conf.DbName = selDb?.EnName ?? ""; 

                            cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                            
                            if (selDb != null && !string.IsNullOrEmpty(selDb.EnName) && _dbMap.ContainsKey(selDb.EnName)) {
                                foreach(var tb in _dbMap[selDb.EnName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            }
                            if (cbTb.Items.Count > 0) cbTb.SelectedIndex = 0;
                        };

                        cbTb.SelectedIndexChanged += (s, ev) => {
                            if (isInitializing) return;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            conf.TableName = selTb?.EnName ?? ""; 

                            if (cbTb.SelectedItem != null && cbDb.SelectedItem != null) {
                                colsLoaded = false; 
                                cbSeg.Text = ""; cbItem.Text = ""; cbLimit.Text = ""; cbDate.Text = ""; cbVal.Text = ""; cbNote.Text = "";
                            }
                        };

                        cbSeg.TextChanged += (s, ev) => { if(!isInitializing) conf.SegCol = cbSeg.Text; };
                        cbItem.TextChanged += (s, ev) => { if(!isInitializing) conf.ItemCol = cbItem.Text; };
                        cbLimit.TextChanged += (s, ev) => { if(!isInitializing) conf.LimitCol = cbLimit.Text; };
                        cbDate.TextChanged += (s, ev) => { if(!isInitializing) conf.DateCol = cbDate.Text; };
                        cbVal.TextChanged += (s, ev) => { if(!isInitializing) conf.ValueCol = cbVal.Text; };
                        cbNote.TextChanged += (s, ev) => { if(!isInitializing) conf.NoteCol = cbNote.Text; };

                        foreach (ItemMap im in cbDb.Items) if (im.EnName == conf.DbName) { cbDb.SelectedItem = im; break; }
                        if (cbDb.SelectedItem != null) {
                            cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                            
                            if (_dbMap.ContainsKey(conf.DbName)) {
                                foreach (var tb in _dbMap[conf.DbName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            }
                            foreach (ItemMap im in cbTb.Items) if (im.EnName == conf.TableName) { cbTb.SelectedItem = im; break; }
                        }
                        
                        cbSeg.Text = conf.SegCol; cbItem.Text = conf.ItemCol; cbLimit.Text = conf.LimitCol;
                        cbDate.Text = conf.DateCol; cbVal.Text = conf.ValueCol; cbNote.Text = conf.NoteCol;

                        isInitializing = false; 

                        pnlRow.Controls.AddRange(new Control[] { btnDel, cbDb, cbTb, cbSeg, cbItem, cbLimit, cbDate, cbVal, cbNote });
                        flpMain.Controls.Add(pnlRow);
                    }

                    Button btnAdd = new Button { Text = "➕ 新增來源", Width = 1250, Height = 45, Margin = new Padding(0, 10, 0, 0), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    btnAdd.Click += (s, ev) => { editingConfigs.Add(new EnvConfigItem()); renderRows(); };
                    
                    flpMain.Controls.Add(btnAdd);

                    flpMain.ResumeLayout(true);
                };

                renderRows();
                pnlScroll.Controls.Add(flpMain);

                Button btnSave = new Button { Text = "💾 儲存設定並重新載入", Dock = DockStyle.Bottom, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSave.FlatAppearance.BorderSize = 0;
                
                btnSave.Click += (senderObj, evnt) => {
                    _configs = editingConfigs;
                    SaveSettings();
                    BtnSearch_Click(null, null);
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(pnlScroll);
                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
        }
    }
}
