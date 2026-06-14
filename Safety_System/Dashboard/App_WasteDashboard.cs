/// FILE: Safety_System/Dashboard/App_WasteDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WasteDashboard
    {
        // 頂部日期區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        // 三大區塊容器
        private DashboardSection _wasteSection;
        private DashboardSection _materialSection;
        private DashboardSection _disposalSection;

        // 🟢 替換：移除 TXT，改為全域物件供 SQLite 取用
        private List<WasteConfigItem> _configs = new List<WasteConfigItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 查詢按鈕，用於防呆禁用
        private Button _btnSearch;

        // 定義下拉選單對應的中英文模型
        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        // 定義自訂統計項目的資料結構
        private class WasteConfigItem
        {
            public string DisplayName { get; set; }
            public string Unit { get; set; } = "噸"; 
            public string Category { get; set; } = "廢棄物產出"; 
            public List<DataSourceDef> Sources { get; set; } = new List<DataSourceDef>();
        }

        private class DataSourceDef
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string FilterValue { get; set; } 
            public string AggType { get; set; } 
            public string ColName2 { get; set; } = "";
        }

        // 封裝區塊 UI 參照
        private class DashboardSection
        {
            public string Category { get; set; }
            public Panel MainPanel { get; set; }
            public Label LblSub1 { get; set; }
            public Label LblSub2 { get; set; }
            public Label LblSub3 { get; set; }
            public Label LblSub4 { get; set; }
            public FlowLayoutPanel PnlData1 { get; set; }
            public FlowLayoutPanel PnlData2 { get; set; }
            public FlowLayoutPanel PnlData3 { get; set; }
            public FlowLayoutPanel PnlData4 { get; set; }
            public Panel TopBox { get; set; }
            public FlowLayoutPanel FlpBottomBox { get; set; }
            public Dictionary<string, Panel> MonthlyPanels { get; set; } = new Dictionary<string, Panel>();
        }

        // 🟢 新增：資料庫初始化邏輯
        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS [WasteDashboardConfigs] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        DisplayName TEXT, Unit TEXT, Category TEXT, DbName TEXT, TableName TEXT, 
                        ColName TEXT, AggType TEXT, FilterValue TEXT, ColName2 TEXT);";
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
            
            TableLayoutPanel masterLayout = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1, 
                RowCount = 5, 
                Margin = new Padding(0)
            };
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            // ==========================================
            // 第一行：大標題
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "♻️ 廢棄物統計及數據分析看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.SeaGreen, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblTitle);

            // ==========================================
            // 第二行：查詢及功能鍵
            // ==========================================
            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 10, 0, 20), Margin = new Padding(0) };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes();

            int btnHeight = 42;

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(110, btnHeight), BackColor = Color.DarkOliveGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            Button btnPdf = new Button { Text = "📄 選擇並導出 PDF", Size = new Size(180, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            Button btnSetting = new Button { Text = "⚙️ 統計設定", Size = new Size(130, btnHeight), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnSetting.Click += BtnSetting_Click;

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 10, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 10, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _btnSearch, btnPdf, btnSetting
            });

            // ==========================================
            // 建置三大區塊：廢棄物產出、原物料產出 與 廢棄物清除
            // ==========================================
            _wasteSection = BuildSection("廢棄物產出數據統計", "廢棄物產出", Color.SeaGreen);
            _materialSection = BuildSection("原物料產出數據統計", "原物料產出", Color.Sienna);
            _disposalSection = BuildSection("廢棄物清除數據統計", "廢棄物清除", Color.SlateGray); 

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);
            masterLayout.Controls.Add(_wasteSection.MainPanel, 0, 2);
            masterLayout.Controls.Add(_materialSection.MainPanel, 0, 3);
            masterLayout.Controls.Add(_disposalSection.MainPanel, 0, 4); 

            mainScrollPanel.Controls.Add(masterLayout);

            _ = LoadDashboardDataAsync();

            return mainScrollPanel;
        }

        private DashboardSection BuildSection(string titleText, string category, Color themeColor)
        {
            var sec = new DashboardSection { Category = category };

            sec.MainPanel = new Panel { Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 0, 0, 30) };

            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // 標題
            Label lblTitle = new Label { 
                Text = $"■ {titleText}", 
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), 
                ForeColor = themeColor, 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.BottomLeft 
            };
            layout.Controls.Add(lblTitle, 0, 0);

            // 第一部份：四大區塊
            sec.TopBox = new Panel { Dock = DockStyle.Fill, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 10, 0, 20) };
            sec.TopBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, sec.TopBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel gridFour = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, RowCount = 2, Padding = new Padding(10) };
            for (int i = 0; i < 4; i++) gridFour.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            sec.LblSub1 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            sec.LblSub2 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            sec.LblSub3 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            sec.LblSub4 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };

            gridFour.Controls.Add(sec.LblSub1, 0, 0); gridFour.Controls.Add(sec.LblSub2, 1, 0);
            gridFour.Controls.Add(sec.LblSub3, 2, 0); gridFour.Controls.Add(sec.LblSub4, 3, 0);

            sec.PnlData1 = CreateDataPanel(); sec.PnlData2 = CreateDataPanel();
            sec.PnlData3 = CreateDataPanel(); sec.PnlData4 = CreateDataPanel();

            gridFour.Controls.Add(sec.PnlData1, 0, 1); gridFour.Controls.Add(sec.PnlData2, 1, 1);
            gridFour.Controls.Add(sec.PnlData3, 2, 1); gridFour.Controls.Add(sec.PnlData4, 3, 1);

            sec.TopBox.Controls.Add(gridFour);
            layout.Controls.Add(sec.TopBox, 0, 1);

            // 第二部份：近三年逐月統計容器
            sec.FlpBottomBox = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0) };
            layout.Controls.Add(sec.FlpBottomBox, 0, 2);

            sec.MainPanel.Controls.Add(layout);
            return sec;
        }

        private FlowLayoutPanel CreateDataPanel()
        {
            return new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 150), FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.FromArgb(248, 249, 250), Margin = new Padding(2), Padding = new Padding(10) };
        }

        private DateTime? ParseUniversalDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return null;
            dateStr = dateStr.Trim().Replace("/", "-");

            Regex twRegex = new Regex(@"^(?<year>\d{2,3})-(?<month>\d{1,2})-(?<day>\d{1,2})(?:\s+.*)?$");
            Match matchTw = twRegex.Match(dateStr);

            if (matchTw.Success)
            {
                if (int.TryParse(matchTw.Groups["year"].Value, out int twYear))
                {
                    if (twYear < 200)
                    {
                        int westernYear = twYear + 1911;
                        string m = matchTw.Groups["month"].Value.PadLeft(2, '0');
                        string d = matchTw.Groups["day"].Value.PadLeft(2, '0');
                        dateStr = $"{westernYear}-{m}-{d}"; 
                    }
                }
            }

            if (DateTime.TryParse(dateStr, out DateTime result)) return result;
            return null;
        }

        private void InitDateComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) { 
                _cboStartYear.Items.Add(i); _cboEndYear.Items.Add(i); 
            }
            for (int i = 1; i <= 12; i++) { 
                _cboStartMonth.Items.Add(i.ToString("D2")); _cboEndMonth.Items.Add(i.ToString("D2")); 
            }
            
            _cboStartYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboStartMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboEndYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            _cboEndMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            
            DateTime today = DateTime.Today;
            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, new DateTime(today.Year, 1, 1));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, today);
        }

        private void UpdateDaysCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null) return;
            int days = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            string currentDay = d.SelectedItem?.ToString();
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));
            if (currentDay != null && d.Items.Contains(currentDay)) d.SelectedItem = currentDay;
            else d.SelectedIndex = d.Items.Count - 1; 
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date)
        {
            y.SelectedItem = date.Year; 
            m.SelectedItem = date.Month.ToString("D2");
            UpdateDaysCombo(y, m, d); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetDateFromCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null || d.SelectedItem == null) return DateTime.Today;
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            return new DateTime((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()), day > maxDay ? maxDay : day);
        }

        private async Task LoadDashboardDataAsync()
        {
            if (_btnSearch != null) _btnSearch.Enabled = false;

            try
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
                DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

                await ProcessSectionAsync(_wasteSection, dtS, dtE);
                await ProcessSectionAsync(_materialSection, dtS, dtE);
                await ProcessSectionAsync(_disposalSection, dtS, dtE);
            }
            finally
            {
                if (_btnSearch != null) _btnSearch.Enabled = true;
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private async Task ProcessSectionAsync(DashboardSection section, DateTime dtS, DateTime dtE)
        {
            var sectionConfigs = _configs.Where(c => c.Category == section.Category).ToList();

            if (sectionConfigs.Count == 0)
            {
                section.PnlData1.Controls.Clear(); section.PnlData2.Controls.Clear(); 
                section.PnlData3.Controls.Clear(); section.PnlData4.Controls.Clear();
                section.PnlData1.Controls.Add(new Label { Text = "此區塊目前無統計項目設定", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                section.FlpBottomBox.Controls.Clear();
                section.MonthlyPanels.Clear();
                return;
            }

            section.LblSub1.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n區間統計總計";
            section.LblSub2.Text = $"【{dtS.AddYears(-1):yyyy/MM/dd} ~ {dtE.AddYears(-1):yyyy/MM/dd}】\n去年同期統計總計";
            section.LblSub3.Text = $"【{dtS.AddYears(-2):yyyy/MM/dd} ~ {dtE.AddYears(-2):yyyy/MM/dd}】\n前年同期統計總計";
            section.LblSub4.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n與去年同期差異分析";

            Dictionary<string, double> dictCurr = new Dictionary<string, double>();
            Dictionary<string, double> dictLy = new Dictionary<string, double>();
            Dictionary<string, double> dictL2y = new Dictionary<string, double>();
            Dictionary<string, DataTable> monthlyTables = new Dictionary<string, DataTable>();

            await Task.Run(() =>
            {
                dictCurr = CalculatePeriodStats(dtS, dtE, sectionConfigs);
                dictLy = CalculatePeriodStats(dtS.AddYears(-1), dtE.AddYears(-1), sectionConfigs);
                dictL2y = CalculatePeriodStats(dtS.AddYears(-2), dtE.AddYears(-2), sectionConfigs);

                int baseYear = dtE.Year;
                int[] years = { baseYear, baseYear - 1, baseYear - 2 };

                foreach (var cfg in sectionConfigs)
                {
                    if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                    DataTable dtMonthly = new DataTable();
                    dtMonthly.Columns.Add("年度", typeof(string));
                    for (int i = 1; i <= 12; i++) dtMonthly.Columns.Add($"{i}月", typeof(double));
                    dtMonthly.Columns.Add("年度總計", typeof(double));

                    foreach (int y in years)
                    {
                        DataRow row = dtMonthly.NewRow();
                        row["年度"] = y.ToString() + "年";
                        double yearlyTotal = 0;

                        for (int m = 1; m <= 12; m++)
                        {
                            DateTime mStart = new DateTime(y, m, 1);
                            DateTime mEnd = new DateTime(y, m, DateTime.DaysInMonth(y, m));
                            
                            var mResult = CalculatePeriodStats(mStart, mEnd, new List<WasteConfigItem> { cfg });
                            double mVal = mResult.ContainsKey(cfg.DisplayName) ? mResult[cfg.DisplayName] : 0;
                            
                            row[$"{m}月"] = mVal;
                            yearlyTotal += mVal;
                        }
                        row["年度總計"] = yearlyTotal;
                        dtMonthly.Rows.Add(row);
                    }
                    monthlyTables[cfg.DisplayName] = dtMonthly;
                }
            });

            // UI 更新 (上半部區塊)
            section.PnlData1.Controls.Clear(); section.PnlData2.Controls.Clear(); 
            section.PnlData3.Controls.Clear(); section.PnlData4.Controls.Clear();

            foreach (var cfg in sectionConfigs)
            {
                if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                string key = cfg.DisplayName;
                string unit = string.IsNullOrEmpty(cfg.Unit) ? "噸" : cfg.Unit;
                string primaryAgg = cfg.Sources.FirstOrDefault()?.AggType ?? "COUNT";
                string format = (primaryAgg == "COUNT") ? "N0" : "N2"; 

                double vCurr = dictCurr.ContainsKey(key) ? dictCurr[key] : 0;
                double vLy = dictLy.ContainsKey(key) ? dictLy[key] : 0;
                double vL2y = dictL2y.ContainsKey(key) ? dictL2y[key] : 0;

                section.PnlData1.Controls.Add(CreateStatLabel(key, vCurr, unit, format));
                section.PnlData2.Controls.Add(CreateStatLabel(key, vLy, unit, format));
                section.PnlData3.Controls.Add(CreateStatLabel(key, vL2y, unit, format));

                double diff = vCurr - vLy;
                string diffText = (diff > 0 ? "+" : "") + diff.ToString(format) + " " + unit;
                Color diffColor = diff > 0 ? Color.IndianRed : (diff < 0 ? Color.ForestGreen : Color.DimGray);

                section.PnlData4.Controls.Add(new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) });
            }

            // UI 更新 (下半部逐月表)
            section.FlpBottomBox.Controls.Clear();
            section.MonthlyPanels.Clear();

            Color headerColor = section.Category == "廢棄物產出" ? Color.SeaGreen : 
                               (section.Category == "原物料產出" ? Color.Sienna : Color.SlateGray);

            foreach (var kvp in monthlyTables)
            {
                string statName = kvp.Key;
                DataTable dt = kvp.Value;
                if (dt == null) continue;

                var matchCfg = sectionConfigs.FirstOrDefault(c => c.DisplayName == statName);
                string primaryAgg = matchCfg?.Sources.FirstOrDefault()?.AggType ?? "COUNT";
                string format = (primaryAgg == "COUNT") ? "N0" : "N2";

                int targetWidth = 1000; 
                if (section.FlpBottomBox.Width > 40) targetWidth = section.FlpBottomBox.Width - 20;
                else if (section.FlpBottomBox.Parent != null && section.FlpBottomBox.Parent.Width > 40) targetWidth = section.FlpBottomBox.Parent.Width - 40;

                Panel pnlWrapper = new Panel { Width = targetWidth, Height = 220, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
                pnlWrapper.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlWrapper.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                
                Label lblTitle = new Label { Text = $"📊 近三年逐月統計：{statName}", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = headerColor, AutoSize = true, Padding = new Padding(15, 10, 0, 10), Dock = DockStyle.Top };
                
                DataGridView dgv = new DataGridView { 
                    Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                    BorderStyle = BorderStyle.None, Margin = new Padding(10)
                };
                
                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = headerColor;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv.ColumnHeadersHeight = 35;
                dgv.RowTemplate.Height = 33;
                dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 250, 245);
                
                dgv.DataSource = dt;

                if (dgv.Columns.Contains("年度")) dgv.Columns["年度"].Width = 100;
                
                if (dgv.Columns.Contains("年度總計")) {
                    dgv.Columns["年度總計"].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                    dgv.Columns["年度總計"].DefaultCellStyle.BackColor = Color.LightYellow;
                    dgv.Columns["年度總計"].DefaultCellStyle.Format = format;
                }
                
                for (int i = 1; i <= 12; i++) {
                    string monthCol = $"{i}月";
                    if (dgv.Columns.Contains(monthCol)) dgv.Columns[monthCol].DefaultCellStyle.Format = format;
                }
                
                dgv.ClearSelection();

                pnlWrapper.Controls.Add(dgv);
                pnlWrapper.Controls.Add(lblTitle);
                
                section.FlpBottomBox.Controls.Add(pnlWrapper);
                section.MonthlyPanels[statName] = pnlWrapper;

                section.FlpBottomBox.Resize += (s, e) => { 
                    int newW = section.FlpBottomBox.ClientSize.Width - 20;
                    if (newW > 0) pnlWrapper.Width = newW; 
                };
            }
        }

        private Label CreateStatLabel(string title, double value, string unit, string format)
        {
            return new Label { Text = $"{title}: {value.ToString(format)} {unit}", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        private Dictionary<string, double> CalculatePeriodStats(DateTime sDate, DateTime eDate, List<WasteConfigItem> targetConfigs)
        {
            var result = new Dictionary<string, double>();
            string sStr = sDate.ToString("yyyy-MM-dd");
            string eStr = eDate.ToString("yyyy-MM-dd");

            foreach (var cfg in targetConfigs)
            {
                if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;
                
                double countVal = 0;
                List<double> gatheredValues = new List<double>();
                string primaryAgg = cfg.Sources.FirstOrDefault()?.AggType ?? "COUNT";

                foreach (var src in cfg.Sources)
                {
                    if (src == null || string.IsNullOrEmpty(src.DbName) || string.IsNullOrEmpty(src.TableName)) continue;

                    try
                    {
                        var cols = DataManager.GetColumnNames(src.DbName, src.TableName);
                        if (cols == null || cols.Count == 0) continue; 

                        string dateCol = cols.Contains("清運日期") ? "清運日期" : (cols.Contains("日期") ? "日期" : (cols.Contains("年月") ? "年月" : (cols.Contains("年度") ? "年度" : "")));
                        if (string.IsNullOrEmpty(dateCol)) continue;

                        DataTable dt = DataManager.GetTableData(src.DbName, src.TableName, dateCol, sStr, eStr);
                        if (dt == null) continue;

                        string filterTarget = string.IsNullOrEmpty(src.FilterValue) || src.FilterValue == "非空值 (有輸入即算)" ? "" : src.FilterValue.Trim();

                        foreach (DataRow r in dt.Rows)
                        {
                            if (r.RowState == DataRowState.Deleted) continue;
                            
                            bool match = false;
                            
                            if (!string.IsNullOrEmpty(src.ColName) && src.ColName != "Id (無條件計數)" && dt.Columns.Contains(src.ColName))
                            {
                                string filterValStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(filterTarget)) {
                                    match = !string.IsNullOrEmpty(filterValStr);
                                } else {
                                    match = filterValStr.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                        .Any(x => x.Trim().Equals(filterTarget, StringComparison.OrdinalIgnoreCase));
                                }
                            }
                            else
                            {
                                match = true; 
                            }

                            if (match)
                            {
                                if (src.AggType == "COUNT")
                                {
                                    countVal++;
                                }
                                else if (src.AggType == "SUM")
                                {
                                    string valStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                    if (double.TryParse(valStr.Replace(",", ""), out double v)) gatheredValues.Add(v);
                                }
                                else if (src.AggType.StartsWith("DIFF"))
                                {
                                    string endStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                    string startStr = r[src.ColName2]?.ToString()?.Trim() ?? "";
                                    
                                    DateTime? endD = ParseUniversalDate(endStr);
                                    DateTime? startD = ParseUniversalDate(startStr);
                                    
                                    if (endD.HasValue && startD.HasValue)
                                    {
                                        double days = (endD.Value.Date - startD.Value.Date).TotalDays;
                                        if (days >= 0) gatheredValues.Add(days);
                                    }
                                }
                            }
                        }
                    }
                    catch { } 
                }

                double finalVal = 0;
                if (primaryAgg == "COUNT") finalVal = countVal;
                else if (primaryAgg == "SUM" || primaryAgg == "DIFF_SUM") finalVal = gatheredValues.Sum();
                else if (primaryAgg == "DIFF_AVG") finalVal = gatheredValues.Any() ? Math.Round(gatheredValues.Average(), 1) : 0;

                result[cfg.DisplayName] = finalVal;
            }
            return result;
        }

        // =========================================================
        // 🟢 設定檔管理與動態設定視窗 (改為 SQLite)
        // =========================================================
        private void LoadSettings()
        {
            _configs.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM WasteDashboardConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        var dict = new Dictionary<string, WasteConfigItem>();
                        while (reader.Read()) {
                            string dispName = reader["DisplayName"].ToString();
                            string unit = reader["Unit"].ToString();
                            string cat = reader["Category"].ToString();
                            
                            if (!dict.ContainsKey(dispName)) {
                                dict[dispName] = new WasteConfigItem { DisplayName = dispName, Unit = unit, Category = cat, Sources = new List<DataSourceDef>() };
                            }

                            dict[dispName].Sources.Add(new DataSourceDef {
                                DbName = reader["DbName"].ToString(),
                                TableName = reader["TableName"].ToString(),
                                ColName = reader["ColName"].ToString(),
                                AggType = reader["AggType"].ToString(),
                                FilterValue = reader["FilterValue"].ToString(),
                                ColName2 = reader["ColName2"].ToString()
                            });
                        }
                        _configs = dict.Values.ToList();
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
                        new SQLiteCommand("DELETE FROM WasteDashboardConfigs", conn, trans).ExecuteNonQuery();
                        
                        string sql = "INSERT INTO WasteDashboardConfigs (DisplayName, Unit, Category, DbName, TableName, ColName, AggType, FilterValue, ColName2) VALUES (@D, @U, @Cat, @DB, @TB, @C, @A, @F, @C2)";
                        foreach (var cfg in _configs) {
                            if (string.IsNullOrEmpty(cfg.DisplayName)) continue;
                            foreach (var src in cfg.Sources) {
                                using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                    cmd.Parameters.AddWithValue("@D", cfg.DisplayName);
                                    cmd.Parameters.AddWithValue("@U", cfg.Unit ?? "噸");
                                    cmd.Parameters.AddWithValue("@Cat", cfg.Category ?? "廢棄物產出");
                                    cmd.Parameters.AddWithValue("@DB", src.DbName ?? "");
                                    cmd.Parameters.AddWithValue("@TB", src.TableName ?? "");
                                    cmd.Parameters.AddWithValue("@C", src.ColName ?? "");
                                    cmd.Parameters.AddWithValue("@A", src.AggType ?? "COUNT");
                                    cmd.Parameters.AddWithValue("@F", src.FilterValue ?? "");
                                    cmd.Parameters.AddWithValue("@C2", src.ColName2 ?? "");
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }

        private void BtnSetting_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "⚙️ 看板自訂統計項目設定", Size = new Size(1450, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                Label l1 = new Label { Text = "已建立的統計項目", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 30 };
                ListBox lbItems = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                pnlLeft.Controls.Add(lbItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnDel);

                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯 / 新增項目", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SeaGreen, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                Panel pName = new Panel { Width = 1100, Height = 45 };
                
                pName.Controls.Add(new Label { Text = "看板分類：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                ComboBox cboCategory = new ComboBox { Width = 140, Location = new Point(100, 7), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                cboCategory.Items.AddRange(new string[] { "廢棄物產出", "原物料產出", "廢棄物清除" });
                cboCategory.SelectedIndex = 0;
                pName.Controls.Add(cboCategory);

                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(250, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 230, Location = new Point(350, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);

                pName.Controls.Add(new Label { Text = "單位：", AutoSize = true, Location = new Point(590, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                TextBox txtUnit = new TextBox { Width = 80, Location = new Point(650, 7), Font = new Font("Microsoft JhengHei UI", 12F), Text = "噸" };
                pName.Controls.Add(txtUnit);
                
                Button btnAddSource = new Button { Text = "➕ 新增項目", Location = new Point(750, 5), Size = new Size(130, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnAddSource.FlatAppearance.BorderSize = 0;
                pName.Controls.Add(btnAddSource);
                
                flpEditor.Controls.Add(pName);

                FlowLayoutPanel flpSourcesContainer = new FlowLayoutPanel { Width = 1150, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                flpEditor.Controls.Add(flpSourcesContainer);

                var editingConfigs = new List<WasteConfigItem>(_configs);

                Action renderRows = null;
                renderRows = () => {
                    flpSourcesContainer.SuspendLayout();
                    flpSourcesContainer.Controls.Clear();

                    var targetConf = editingConfigs.FirstOrDefault(c => c.DisplayName == txtName.Text);
                    if (targetConf == null) {
                        targetConf = new WasteConfigItem();
                        editingConfigs.Add(targetConf);
                    }

                    for (int i = 0; i < targetConf.Sources.Count; i++) {
                        int currentIndex = i;
                        var srcDef = targetConf.Sources[i];

                        Panel pRow = new Panel { Width = 1100, Height = 75, BackColor = Color.FromArgb(245, 250, 245), Margin = new Padding(0, 5, 0, 5) };
                        pRow.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pRow.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                        
                        int ly = 10, cy = 35;
                        int x0 = 10, w0 = 35;   
                        int x1 = 55, w1 = 110;  
                        int x2 = 175, w2 = 140; 
                        int x3 = 325, w3 = 145; 
                        int x4 = 480, w4 = 155; 
                        int x5 = 645, w5 = 145; 
                        int x6 = 800, w6 = 140; 
                        int x7 = 950, w7 = 120; 

                        Button btnRemove = new Button { Text = "❌", Location = new Point(x0, 34), Width = w0, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        btnRemove.FlatAppearance.BorderSize = 0;
                        btnRemove.Click += (s, ev) => { targetConf.Sources.RemoveAt(currentIndex); renderRows(); };

                        ComboBox cbDb = new ComboBox { Location = new Point(x1, cy), Width = w1, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbTb = new ComboBox { Location = new Point(x2, cy), Width = w2, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbCol = new ComboBox { Location = new Point(x3, cy), Width = w3, DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbFilter = new ComboBox { Location = new Point(x4, cy), Width = w4, DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbAgg = new ComboBox { Location = new Point(x5, cy), Width = w5, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        
                        Label lblColEnd = new Label { Text = "主欄位(迄)", Location = new Point(x6, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) };
                        ComboBox cbColEnd = new ComboBox { Location = new Point(x6, cy), Width = w6, DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F), Visible = false };
                        
                        Label lblColStart = new Label { Text = "次欄位(起)", Location = new Point(x7, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), ForeColor = Color.DarkOrange, Visible = false };
                        ComboBox cbColStart = new ComboBox { Location = new Point(x7, cy), Width = w7, DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F), Visible = false };

                        pRow.Controls.AddRange(new Control[] {
                            btnRemove, 
                            new Label { Text = "資料庫", Location = new Point(x1, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbDb,
                            new Label { Text = "資料表", Location = new Point(x2, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbTb,
                            new Label { Text = "計算欄位", Location = new Point(x3, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbCol,
                            new Label { Text = "選項過濾條件", Location = new Point(x4, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbFilter,
                            new Label { Text = "運算方式", Location = new Point(x5, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbAgg,
                            lblColEnd, cbColEnd,
                            lblColStart, cbColStart
                        });

                        cbDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                        foreach (var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

                        bool isInitializing = true;
                        bool colsLoaded = false;

                        Action lazyLoadCols = () => {
                            if (colsLoaded) return;
                            cbCol.Items.Clear(); cbCol.Items.Add("Id (無條件計數)");
                            cbColEnd.Items.Clear(); cbColEnd.Items.Add("");
                            cbColStart.Items.Clear(); cbColStart.Items.Add("");

                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName)) {
                                var cols = DataManager.GetColumnNames(selDb.EnName, selTb.EnName).Where(c => c != "Id");
                                foreach (var c in cols) {
                                    cbCol.Items.Add(c);
                                    cbColEnd.Items.Add(c);
                                    cbColStart.Items.Add(c);
                                }
                            }
                            colsLoaded = true;
                        };

                        EventHandler triggerLoad = (s, ev) => { lazyLoadCols(); };
                        cbCol.DropDown += triggerLoad;
                        cbColEnd.DropDown += triggerLoad;
                        cbColStart.DropDown += triggerLoad;
                        cbFilter.DropDown += triggerLoad;

                        cbDb.SelectedIndexChanged += (s, ev) => {
                            if (isInitializing) return;
                            var selDb = cbDb.SelectedItem as ItemMap;
                            srcDef.DbName = selDb?.EnName ?? "";

                            cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                            if (selDb != null && !string.IsNullOrEmpty(selDb.EnName) && _dbMap.ContainsKey(selDb.EnName)) {
                                foreach (var tb in _dbMap[selDb.EnName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            }
                            if (cbTb.Items.Count > 0) cbTb.SelectedIndex = 0;
                        };

                        cbTb.SelectedIndexChanged += (s, ev) => {
                            if (isInitializing) return;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            srcDef.TableName = selTb?.EnName ?? "";

                            if (cbTb.SelectedItem != null && cbDb.SelectedItem != null) {
                                colsLoaded = false;
                                cbCol.Text = "Id (無條件計數)"; cbFilter.Text = "非空值 (有輸入即算)";
                            }
                        };

                        cbAgg.Items.AddRange(new string[] { "計數", "加總", "日期相減(總天數)" });
                        cbAgg.SelectedIndexChanged += (s, ev) => {
                            bool isDiff = cbAgg.Text.Contains("相減");
                            lblColEnd.Visible = isDiff; cbColEnd.Visible = isDiff;
                            lblColStart.Visible = isDiff; cbColStart.Visible = isDiff;
                            if(!isInitializing) {
                                if (cbAgg.Text == "加總") srcDef.AggType = "SUM";
                                else if (cbAgg.Text == "日期相減(總天數)") srcDef.AggType = "DIFF_SUM";
                                else srcDef.AggType = "COUNT";
                            }
                        };

                        cbCol.SelectedIndexChanged += (s, ev) => {
                            cbFilter.Items.Clear(); 
                            cbFilter.Items.Add("非空值 (有輸入即算)");
                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            string col = cbCol.Text;
                            
                            if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName) && col != "Id (無條件計數)") {
                                string tbName = selTb.EnName;
                                string multiKey = $"{tbName}|{col}";

                                if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey)) {
                                    foreach (var opt in App_DropdownManager.MultiSelectCache[multiKey]) {
                                        if (!string.IsNullOrWhiteSpace(opt.Text) && !cbFilter.Items.Contains(opt.Text.Trim())) {
                                            cbFilter.Items.Add(opt.Text.Trim());
                                        }
                                    }
                                }
                                
                                string[] dropOpts = App_DropdownManager.GetAllOptionsForColumn(tbName, col);
                                if (dropOpts != null && dropOpts.Length > 0) {
                                    foreach (var opt in dropOpts) {
                                        if (!string.IsNullOrWhiteSpace(opt) && !cbFilter.Items.Contains(opt.Trim())) {
                                            cbFilter.Items.Add(opt.Trim());
                                        }
                                    }
                                }
                            }
                            cbFilter.SelectedIndex = 0;
                        };

                        // 值改變時即時更新回模型
                        cbCol.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.ColName = cbCol.Text; } };
                        cbFilter.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.FilterValue = cbFilter.Text; } };
                        cbColEnd.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.ColName = cbColEnd.Text; } };
                        cbColStart.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.ColName2 = cbColStart.Text; } };

                        // 🟢 初始化
                        foreach (ItemMap im in cbDb.Items) if (im.EnName == srcDef.DbName) { cbDb.SelectedItem = im; break; }
                        if (cbDb.SelectedItem != null && _dbMap.ContainsKey(srcDef.DbName)) {
                            cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                            foreach (var tb in _dbMap[srcDef.DbName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            foreach (ItemMap im in cbTb.Items) if (im.EnName == srcDef.TableName) { cbTb.SelectedItem = im; break; }
                        }

                        cbCol.Text = srcDef.ColName;
                        if (!string.IsNullOrEmpty(srcDef.ColName2)) {
                            cbColEnd.Text = srcDef.ColName;
                            cbColStart.Text = srcDef.ColName2;
                        }

                        if (!string.IsNullOrEmpty(srcDef.FilterValue)) {
                            if (!cbFilter.Items.Contains(srcDef.FilterValue)) cbFilter.Items.Add(srcDef.FilterValue);
                            cbFilter.Text = srcDef.FilterValue;
                        } else {
                            cbFilter.Text = "非空值 (有輸入即算)";
                        }
                        
                        if (srcDef.AggType == "COUNT") cbAgg.Text = "計數";
                        else if (srcDef.AggType == "SUM") cbAgg.Text = "加總";
                        else if (srcDef.AggType == "DIFF_SUM") cbAgg.Text = "日期相減(總天數)";
                        else cbAgg.Text = "計數"; 

                        isInitializing = false;
                        flpSourcesContainer.Controls.Add(pRow);
                    }

                    Button btnAdd = new Button { Text = "➕ 新增來源", Width = 1100, Height = 45, Margin = new Padding(0, 10, 0, 0), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    btnAdd.Click += (s, ev) => { 
                        var targetConf = editingConfigs.FirstOrDefault(c => c.DisplayName == txtName.Text);
                        if (targetConf == null) {
                            targetConf = new WasteConfigItem { DisplayName = txtName.Text, Unit = txtUnit.Text, Category = cboCategory.Text };
                            editingConfigs.Add(targetConf);
                        }
                        targetConf.Sources.Add(new DataSourceDef()); 
                        renderRows(); 
                    };
                    
                    flpSourcesContainer.Controls.Add(btnAdd);
                    flpSourcesContainer.ResumeLayout(true);
                };

                renderRows();

                Button btnSaveRow = new Button { Text = "💾 儲存並加入清單", Width = 900, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 0), Cursor = Cursors.Hand };

                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);
                pnlRight.Controls.Add(btnSaveRow);
                btnSaveRow.Dock = DockStyle.Bottom;

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);
                f.Controls.Add(tlp);

                Action refreshList = () => {
                    lbItems.Items.Clear();
                    foreach (var cfg in editingConfigs) {
                        if (cfg != null && !string.IsNullOrEmpty(cfg.DisplayName)) {
                            lbItems.Items.Add($"[{cfg.Category}] {cfg.DisplayName}");
                        }
                    }
                };
                refreshList();

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (lbItems.SelectedIndex < 0) return;
                    flpSourcesContainer.Controls.Clear();
                    var cfg = editingConfigs[lbItems.SelectedIndex];
                    
                    cboCategory.SelectedItem = cfg.Category;
                    txtName.Text = cfg.DisplayName;
                    txtUnit.Text = string.IsNullOrEmpty(cfg.Unit) ? "噸" : cfg.Unit; 
                    
                    renderRows();
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        editingConfigs.RemoveAt(lbItems.SelectedIndex);
                        refreshList();
                        txtName.Clear();
                        flpSourcesContainer.Controls.Clear();
                    }
                };

                btnSaveRow.Click += (ss, ee) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入顯示名稱！"); return; }
                    
                    var newCfg = new WasteConfigItem { 
                        DisplayName = txtName.Text.Trim(),
                        Unit = string.IsNullOrWhiteSpace(txtUnit.Text) ? "噸" : txtUnit.Text.Trim(),
                        Category = cboCategory.SelectedItem.ToString()
                    };
                    
                    var oldCfg = editingConfigs.FirstOrDefault(c => c.DisplayName == txtName.Text);
                    if (oldCfg != null) {
                        newCfg.Sources = oldCfg.Sources; 
                        int idx = editingConfigs.IndexOf(oldCfg);
                        editingConfigs[idx] = newCfg;
                    } else {
                        editingConfigs.Add(newCfg);
                    }

                    _configs = new List<WasteConfigItem>(editingConfigs);
                    SaveSettings();
                    refreshList();
                    MessageBox.Show("儲存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                f.ShowDialog();
                _ = LoadDashboardDataAsync(); 
            }
        }

        // =========================================================
        // PDF 導出 
        // =========================================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 550, Height = 600, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));

                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Fill, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), AutoSize = true };
                tlp.Controls.Add(lbl, 0, 0);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(15, 5, 15, 5), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
                
                clb.Items.Add("[廢棄物產出] 區間統計總計 (四大區塊)", true); 
                foreach (var kvp in _wasteSection.MonthlyPanels) {
                    clb.Items.Add($"[廢棄物產出] 近三年逐月：{kvp.Key}", true);
                }

                clb.Items.Add("[原物料產出] 區間統計總計 (四大區塊)", true); 
                foreach (var kvp in _materialSection.MonthlyPanels) {
                    clb.Items.Add($"[原物料產出] 近三年逐月：{kvp.Key}", true);
                }

                clb.Items.Add("[廢棄物清除] 區間統計總計 (四大區塊)", true); 
                foreach (var kvp in _disposalSection.MonthlyPanels) {
                    clb.Items.Add($"[廢棄物清除] 近三年逐月：{kvp.Key}", true);
                }
                
                tlp.Controls.Add(clb, 0, 1);

                Panel pnlBottom = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) };
                
                Button btnSelectAll = new Button { Text = "☑️ 全選", Location = new Point(15, 5), Size = new Size(100, 35), BackColor = Color.LightGray, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button btnUnselectAll = new Button { Text = "☐ 取消全選", Location = new Point(125, 5), Size = new Size(130, 35), BackColor = Color.LightGray, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 40, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                
                btnSelectAll.Click += (s, e) => {
                    for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true);
                };

                btnUnselectAll.Click += (s, e) => {
                    for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false);
                };

                pnlBottom.Controls.Add(btnSelectAll);
                pnlBottom.Controls.Add(btnUnselectAll);
                pnlBottom.Controls.Add(btnOk);
                
                tlp.Controls.Add(pnlBottom, 0, 2);

                f.Controls.Add(tlp);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    foreach (var item in clb.CheckedItems) {
                        string text = item.ToString();
                        
                        if (text == "[廢棄物產出] 區間統計總計 (四大區塊)") {
                            selectedPanels.Add(_wasteSection.TopBox);
                        } else if (text == "[原物料產出] 區間統計總計 (四大區塊)") {
                            selectedPanels.Add(_materialSection.TopBox);
                        } else if (text == "[廢棄物清除] 區間統計總計 (四大區塊)") {
                            selectedPanels.Add(_disposalSection.TopBox);
                        } else if (text.StartsWith("[廢棄物產出] 近三年逐月：")) {
                            string key = text.Replace("[廢棄物產出] 近三年逐月：", "");
                            if (_wasteSection.MonthlyPanels.ContainsKey(key)) selectedPanels.Add(_wasteSection.MonthlyPanels[key]);
                        } else if (text.StartsWith("[原物料產出] 近三年逐月：")) {
                            string key = text.Replace("[原物料產出] 近三年逐月：", "");
                            if (_materialSection.MonthlyPanels.ContainsKey(key)) selectedPanels.Add(_materialSection.MonthlyPanels[key]);
                        } else if (text.StartsWith("[廢棄物清除] 近三年逐月：")) {
                            string key = text.Replace("[廢棄物清除] 近三年逐月：", "");
                            if (_disposalSection.MonthlyPanels.ContainsKey(key)) selectedPanels.Add(_disposalSection.MonthlyPanels[key]);
                        }
                    }
                }
            }
            return selectedPanels;
        }

        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (_configs.Count == 0) { MessageBox.Show("無資料可導出。"); return; }

            var panelsToExport = GetSelectedExportPanels();
            if (panelsToExport.Count == 0) return;

            List<Bitmap> bitmaps = new List<Bitmap>();
            foreach (var pnl in panelsToExport) 
            {
                Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                bitmaps.Add(bmp);
            }

            string dateStr = $"查詢區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text} ~ {_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
            
            PdfHelper.ExportDashboardToPdf(bitmaps, "廢棄物統計及數據分析表", dateStr, "廢棄物統計及數據分析表");
        }
    }
}
