/// FILE: Safety_System/Dashboard/App_TestDashboard.cs ///
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

namespace Safety_System
{
    public class App_TestDashboard
    {
        // 頂部日期區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        // 四大區塊的標題與內容容器
        private Label _lblSub1, _lblSub2, _lblSub3, _lblSub4;
        private FlowLayoutPanel _pnlData1, _pnlData2, _pnlData3, _pnlData4;

        // 截圖與匯出用的外框
        private Panel _pnlTopBox;
        private FlowLayoutPanel _flpBottomBox; 

        // 存放動態生成的 Grid 以供 PDF 導出時對應
        private Dictionary<string, Panel> _monthlyPanels = new Dictionary<string, Panel>();

        private List<TestConfigItem> _configs = new List<TestConfigItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        private Button _btnSearch;

        // 目標資料表清單 (11個)
        private readonly string[] _targetTables = { 
            "EnvMonitor", "WastewaterPeriodic", "DrinkingWater", "IndustrialZoneTest", 
            "SoilGasTest", "WastewaterSelfTest", "CoolingWaterVendor", "CoolingWaterSelf", 
            "TCLP", "WaterMeterCalibration", "OtherTests" 
        };

        // 定義下拉選單對應的中英文模型
        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        // 定義自訂統計項目的資料結構
        private class TestConfigItem
        {
            public string DisplayName { get; set; }
            public string Unit { get; set; } = "mg/L"; 
            public List<DataSourceDef> Sources { get; set; } = new List<DataSourceDef>();

            public TestConfigItem Clone() {
                var clone = new TestConfigItem { DisplayName = this.DisplayName, Unit = this.Unit };
                foreach (var s in this.Sources) {
                    clone.Sources.Add(new DataSourceDef {
                        DbName = s.DbName, TableName = s.TableName, ColName = s.ColName,
                        FilterValue = s.FilterValue, AggType = s.AggType, ColName2 = s.ColName2, RefColName = s.RefColName
                    });
                }
                return clone;
            }
        }

        private class DataSourceDef
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string FilterValue { get; set; } 
            public string AggType { get; set; } 
            public string ColName2 { get; set; } = ""; 
            public string RefColName { get; set; } = ""; 
        }

        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS [TestDashboardConfigs] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        DisplayName TEXT, Unit TEXT, DbName TEXT, TableName TEXT, 
                        ColName TEXT, AggType TEXT, FilterValue TEXT, ColName2 TEXT, RefColName TEXT);";
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
                RowCount = 4,
                Margin = new Padding(0)
            };
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            // ==========================================
            // 第一行：大標題
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "📋 檢測數據綜合管理看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.Chocolate, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
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

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(100, btnHeight), BackColor = Color.SaddleBrown, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            Button btnPdf = new Button { Text = "📄 選擇並導出 PDF", Size = new Size(180, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.FlatAppearance.BorderSize = 0;
            btnPdf.Click += BtnPdf_Click;

            Button btnSetting = new Button { Text = "⚙️ 統計設定", Size = new Size(130, btnHeight), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnSetting.FlatAppearance.BorderSize = 0;
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
            // 第三行：四大區塊 (區間統計等)
            // ==========================================
            _pnlTopBox = new Panel { Dock = DockStyle.Fill, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            _pnlTopBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlTopBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel gridFour = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, RowCount = 2, Padding = new Padding(10) };
            for (int i = 0; i < 4; i++) gridFour.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Color headColor = Color.Chocolate;
            _lblSub1 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            _lblSub2 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            _lblSub3 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            _lblSub4 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };

            gridFour.Controls.Add(_lblSub1, 0, 0); gridFour.Controls.Add(_lblSub2, 1, 0);
            gridFour.Controls.Add(_lblSub3, 2, 0); gridFour.Controls.Add(_lblSub4, 3, 0);

            _pnlData1 = CreateDataPanel(); _pnlData2 = CreateDataPanel();
            _pnlData3 = CreateDataPanel(); _pnlData4 = CreateDataPanel();

            gridFour.Controls.Add(_pnlData1, 0, 1); gridFour.Controls.Add(_pnlData2, 1, 1);
            gridFour.Controls.Add(_pnlData3, 2, 1); gridFour.Controls.Add(_pnlData4, 3, 1);

            _pnlTopBox.Controls.Add(gridFour);

            // ==========================================
            // 第四行：近三年逐月統計 (各自獨立表)
            // ==========================================
            _flpBottomBox = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0) };

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);
            masterLayout.Controls.Add(_pnlTopBox, 0, 2);
            masterLayout.Controls.Add(_flpBottomBox, 0, 3);

            mainScrollPanel.Controls.Add(masterLayout);

            _ = LoadDashboardDataAsync();

            return mainScrollPanel;
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

            if (DateTime.TryParse(dateStr, out DateTime result))
            {
                return result;
            }

            return null;
        }

        private void InitDateComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) 
            { 
                _cboStartYear.Items.Add(i); 
                _cboEndYear.Items.Add(i); 
            }
            for (int i = 1; i <= 12; i++) 
            { 
                _cboStartMonth.Items.Add(i.ToString("D2")); 
                _cboEndMonth.Items.Add(i.ToString("D2")); 
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
            int days = DateTime.DaysInMonth(int.Parse(y.SelectedItem.ToString()), int.Parse(m.SelectedItem.ToString()));
            string currentDay = d.SelectedItem?.ToString();
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));
            if (currentDay != null && d.Items.Contains(currentDay)) d.SelectedItem = currentDay;
            else d.SelectedIndex = d.Items.Count - 1; 
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date)
        {
            y.SelectedItem = date.Year.ToString(); 
            m.SelectedItem = date.Month.ToString("D2");
            UpdateDaysCombo(y, m, d); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetDateFromCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null || d.SelectedItem == null) return DateTime.Today;
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth(int.Parse(y.SelectedItem.ToString()), int.Parse(m.SelectedItem.ToString()));
            return new DateTime(int.Parse(y.SelectedItem.ToString()), int.Parse(m.SelectedItem.ToString()), day > maxDay ? maxDay : day);
        }

        private async Task LoadDashboardDataAsync()
        {
            if (_btnSearch != null) _btnSearch.Enabled = false;

            try
            {
                if (_configs.Count == 0)
                {
                    _pnlData1.Controls.Clear(); _pnlData2.Controls.Clear(); _pnlData3.Controls.Clear(); _pnlData4.Controls.Clear();
                    _pnlData1.Controls.Add(new Label { Text = "請先點擊上方 [統計設定]\n新增欲追蹤的指標！", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                    _flpBottomBox.Controls.Clear();
                    _monthlyPanels.Clear();
                    return;
                }

                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
                DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

                _lblSub1.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n區間統計總計";
                _lblSub2.Text = $"【{dtS.AddYears(-1):yyyy/MM/dd} ~ {dtE.AddYears(-1):yyyy/MM/dd}】\n去年同期統計總計";
                _lblSub3.Text = $"【{dtS.AddYears(-2):yyyy/MM/dd} ~ {dtE.AddYears(-2):yyyy/MM/dd}】\n前年同期統計總計";
                _lblSub4.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n與去年同期差異分析";

                Dictionary<string, double> dictCurr = new Dictionary<string, double>();
                Dictionary<string, double> dictLy = new Dictionary<string, double>();
                Dictionary<string, double> dictL2y = new Dictionary<string, double>();

                Dictionary<string, DataTable> monthlyTables = new Dictionary<string, DataTable>();

                await Task.Run(() =>
                {
                    dictCurr = CalculatePeriodStats(dtS, dtE);
                    dictLy = CalculatePeriodStats(dtS.AddYears(-1), dtE.AddYears(-1));
                    dictL2y = CalculatePeriodStats(dtS.AddYears(-2), dtE.AddYears(-2));

                    int baseYear = dtE.Year;
                    int[] years = { baseYear, baseYear - 1, baseYear - 2 };

                    foreach (var cfg in _configs)
                    {
                        if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                        DataTable dtMonthly = new DataTable();
                        dtMonthly.Columns.Add("年度", typeof(string));
                        for (int i = 1; i <= 12; i++) dtMonthly.Columns.Add($"{i}月", typeof(double));
                        
                        dtMonthly.Columns.Add("最大值", typeof(double));
                        dtMonthly.Columns.Add("最小值", typeof(double));
                        dtMonthly.Columns.Add("平均值", typeof(double));

                        foreach (int y in years)
                        {
                            DataRow row = dtMonthly.NewRow();
                            row["年度"] = y.ToString() + "年";
                            List<double> yearlyValidValues = new List<double>();

                            for (int m = 1; m <= 12; m++)
                            {
                                DateTime mStart = new DateTime(y, m, 1);
                                DateTime mEnd = new DateTime(y, m, DateTime.DaysInMonth(y, m));
                                
                                var mResult = CalculatePeriodStats(mStart, mEnd, new List<TestConfigItem> { cfg });
                                double mVal = mResult.ContainsKey(cfg.DisplayName) ? mResult[cfg.DisplayName] : 0;
                                
                                row[$"{m}月"] = mVal;
                                
                                if (mVal != 0) {
                                    yearlyValidValues.Add(mVal); 
                                }
                            }
                            
                            row["最大值"] = yearlyValidValues.Any() ? yearlyValidValues.Max() : 0;
                            row["最小值"] = yearlyValidValues.Any() ? yearlyValidValues.Min() : 0;
                            row["平均值"] = yearlyValidValues.Any() ? Math.Round(yearlyValidValues.Average(), 2) : 0;

                            dtMonthly.Rows.Add(row);
                        }
                        monthlyTables[cfg.DisplayName] = dtMonthly;
                    }
                });

                // ================= UI 更新 (上半部區塊) =================
                _pnlData1.Controls.Clear(); _pnlData2.Controls.Clear(); _pnlData3.Controls.Clear(); _pnlData4.Controls.Clear();
                foreach (var cfg in _configs)
                {
                    if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                    string key = cfg.DisplayName;
                    string unit = string.IsNullOrEmpty(cfg.Unit) ? "mg/L" : cfg.Unit;

                    string primaryAgg = cfg.Sources.FirstOrDefault()?.AggType ?? "COUNT";
                    string format = (primaryAgg == "COUNT") ? "N0" : "N2"; 

                    double vCurr = dictCurr.ContainsKey(key) ? dictCurr[key] : 0;
                    double vLy = dictLy.ContainsKey(key) ? dictLy[key] : 0;
                    double vL2y = dictL2y.ContainsKey(key) ? dictL2y[key] : 0;

                    _pnlData1.Controls.Add(CreateStatLabel(key, vCurr, unit, format));
                    _pnlData2.Controls.Add(CreateStatLabel(key, vLy, unit, format));
                    _pnlData3.Controls.Add(CreateStatLabel(key, vL2y, unit, format));

                    double diff = vCurr - vLy;
                    string diffText = (diff > 0 ? "+" : "") + diff.ToString(format) + " " + unit;
                    Color diffColor = diff > 0 ? Color.IndianRed : (diff < 0 ? Color.ForestGreen : Color.DimGray);

                    _pnlData4.Controls.Add(new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) });
                }

                // ================= UI 更新 (下半部逐月表) =================
                _flpBottomBox.Controls.Clear();
                _monthlyPanels.Clear();

                foreach (var kvp in monthlyTables)
                {
                    try 
                    {
                        string statName = kvp.Key;
                        DataTable dt = kvp.Value;
                        if (dt == null) continue;

                        var matchCfg = _configs.FirstOrDefault(c => c.DisplayName == statName);
                        string primaryAgg = matchCfg?.Sources.FirstOrDefault()?.AggType ?? "COUNT";
                        string format = (primaryAgg == "COUNT") ? "N0" : "N2";

                        int targetWidth = 1000; 
                        if (_flpBottomBox.Width > 40) targetWidth = _flpBottomBox.Width - 20;
                        else if (_flpBottomBox.Parent != null && _flpBottomBox.Parent.Width > 40) targetWidth = _flpBottomBox.Parent.Width - 40;

                        Panel pnlWrapper = new Panel { Width = targetWidth, Height = 220, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
                        pnlWrapper.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlWrapper.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                        
                        Label lblTitle = new Label { Text = $"📊 近三年逐月統計：{statName}", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Chocolate, AutoSize = true, Padding = new Padding(15, 10, 0, 10), Dock = DockStyle.Top };
                        
                        DataGridView dgv = new DataGridView { 
                            Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                            BorderStyle = BorderStyle.None, Margin = new Padding(10)
                        };
                        
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.Chocolate;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dgv.ColumnHeadersHeight = 35;
                        dgv.RowTemplate.Height = 33;
                        dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.SeaShell;
                        
                        dgv.DataSource = dt;

                        if (dgv.Columns.Contains("年度")) {
                            dgv.Columns["年度"].Width = 100;
                        }
                        
                        string[] statCols = { "最大值", "最小值", "平均值" };
                        foreach(var sc in statCols) {
                            if (dgv.Columns.Contains(sc)) {
                                dgv.Columns[sc].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                                dgv.Columns[sc].DefaultCellStyle.BackColor = Color.LightYellow;
                                dgv.Columns[sc].DefaultCellStyle.Format = format;
                            }
                        }
                        
                        for (int i = 1; i <= 12; i++) {
                            string monthCol = $"{i}月";
                            if (dgv.Columns.Contains(monthCol)) {
                                dgv.Columns[monthCol].DefaultCellStyle.Format = format;
                            }
                        }

                        // 事件觸發：遇到 0 則字體變成灰色
                        dgv.CellFormatting += (s, ev) => {
                            if (ev.Value != null && ev.RowIndex >= 0 && ev.ColumnIndex > 0) { 
                                if (double.TryParse(ev.Value.ToString(), out double val)) {
                                    if (val == 0) {
                                        ev.CellStyle.ForeColor = Color.LightGray;
                                    }
                                }
                            }
                        };
                        
                        dgv.ClearSelection();

                        pnlWrapper.Controls.Add(dgv);
                        pnlWrapper.Controls.Add(lblTitle);
                        
                        _flpBottomBox.Controls.Add(pnlWrapper);
                        _monthlyPanels[statName] = pnlWrapper;

                        _flpBottomBox.Resize += (s, e) => { 
                            int newW = _flpBottomBox.ClientSize.Width - 20;
                            if (newW > 0) pnlWrapper.Width = newW; 
                        };
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"繪製表格 {kvp.Key} 時發生錯誤: {ex.Message}");
                    }
                }
            }
            finally
            {
                if (_btnSearch != null) _btnSearch.Enabled = true;
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private Label CreateStatLabel(string title, double value, string unit, string format)
        {
            return new Label { Text = $"{title}: {value.ToString(format)} {unit}", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        private Dictionary<string, double> CalculatePeriodStats(DateTime sDate, DateTime eDate, List<TestConfigItem> targetConfigs = null)
        {
            var result = new Dictionary<string, double>();
            var configsToRun = targetConfigs ?? _configs;
            string sStr = sDate.ToString("yyyy-MM-dd");
            string eStr = eDate.ToString("yyyy-MM-dd");

            foreach (var cfg in configsToRun)
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

                        string dateCol = cols.Contains("日期") ? "日期" : (cols.Contains("年月") ? "年月" : (cols.Contains("年度") ? "年度" : ""));
                        if (string.IsNullOrEmpty(dateCol)) continue;

                        DataTable dt = DataManager.GetTableData(src.DbName, src.TableName, dateCol, sStr, eStr);
                        if (dt == null) continue;

                        string filterTarget = string.IsNullOrEmpty(src.FilterValue) || src.FilterValue == "非空值 (有輸入即算)" ? "" : src.FilterValue.Trim();

                        // 判斷要被當作篩選對象的欄位 (如果參照欄位存在就用參照，否則用計算欄位)
                        string filterColName = !string.IsNullOrEmpty(src.RefColName) ? src.RefColName : src.ColName;

                        foreach (DataRow r in dt.Rows)
                        {
                            if (r.RowState == DataRowState.Deleted) continue;
                            
                            bool match = false;
                            
                            if (!string.IsNullOrEmpty(filterColName) && filterColName != "Id (無條件計數)" && dt.Columns.Contains(filterColName))
                            {
                                string filterValStr = r[filterColName]?.ToString()?.Trim() ?? "";
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
                                else if (src.AggType == "AVG" || src.AggType == "MAX" || src.AggType == "MIN")
                                {
                                    if (src.ColName != "Id (無條件計數)" && dt.Columns.Contains(src.ColName)) {
                                        string valStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                        if (double.TryParse(valStr.Replace(",", ""), out double v)) gatheredValues.Add(v);
                                    }
                                }
                            }
                        }
                    }
                    catch { } 
                }

                double finalVal = 0;
                if (primaryAgg == "COUNT") finalVal = countVal;
                else if (primaryAgg == "AVG") finalVal = gatheredValues.Any() ? Math.Round(gatheredValues.Average(), 2) : 0;
                else if (primaryAgg == "MAX") finalVal = gatheredValues.Any() ? gatheredValues.Max() : 0;
                else if (primaryAgg == "MIN") finalVal = gatheredValues.Any() ? gatheredValues.Min() : 0;

                result[cfg.DisplayName] = finalVal;
            }
            return result;
        }

        private void LoadSettings()
        {
            _configs.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM TestDashboardConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        var dict = new Dictionary<string, TestConfigItem>();
                        while (reader.Read()) {
                            string dispName = reader["DisplayName"].ToString();
                            string unit = reader["Unit"].ToString();
                            
                            if (!dict.ContainsKey(dispName)) {
                                dict[dispName] = new TestConfigItem { DisplayName = dispName, Unit = unit, Sources = new List<DataSourceDef>() };
                            }

                            dict[dispName].Sources.Add(new DataSourceDef {
                                DbName = reader["DbName"].ToString(),
                                TableName = reader["TableName"].ToString(),
                                ColName = reader["ColName"].ToString(),
                                AggType = reader["AggType"].ToString(),
                                FilterValue = reader["FilterValue"].ToString(),
                                ColName2 = reader["ColName2"].ToString(),
                                RefColName = reader["RefColName"].ToString()
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
                        new SQLiteCommand("DELETE FROM TestDashboardConfigs", conn, trans).ExecuteNonQuery();
                        
                        string sql = "INSERT INTO TestDashboardConfigs (DisplayName, Unit, DbName, TableName, ColName, AggType, FilterValue, ColName2, RefColName) VALUES (@D, @U, @DB, @TB, @C, @A, @F, @C2, @RC)";
                        foreach (var cfg in _configs) {
                            if (string.IsNullOrEmpty(cfg.DisplayName)) continue;
                            foreach (var src in cfg.Sources) {
                                using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                    cmd.Parameters.AddWithValue("@D", cfg.DisplayName);
                                    cmd.Parameters.AddWithValue("@U", cfg.Unit ?? "mg/L");
                                    cmd.Parameters.AddWithValue("@DB", src.DbName ?? "");
                                    cmd.Parameters.AddWithValue("@TB", src.TableName ?? "");
                                    cmd.Parameters.AddWithValue("@C", src.ColName ?? "");
                                    cmd.Parameters.AddWithValue("@A", src.AggType ?? "COUNT");
                                    cmd.Parameters.AddWithValue("@F", src.FilterValue ?? "");
                                    cmd.Parameters.AddWithValue("@C2", src.ColName2 ?? "");
                                    cmd.Parameters.AddWithValue("@RC", src.RefColName ?? "");
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
            using (Form f = new Form { Text = "⚙️ 看板自訂統計項目設定", Size = new Size(1300, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                // 🟢 左側面板：清單與新增鈕
                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                
                Button btnAddNewStat = new Button { Text = "➕ 新增左側統計指標", Dock = DockStyle.Top, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
                btnAddNewStat.FlatAppearance.BorderSize = 0;

                Label l1 = new Label { Text = "已建立的統計項目", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 35, Padding = new Padding(0,10,0,0) };
                ListBox lbItems = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                pnlLeft.Controls.Add(lbItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnAddNewStat);
                pnlLeft.Controls.Add(btnDel);

                // 🟢 右側面板：編輯區
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯 / 新增項目", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                Panel pName = new Panel { Width = 950, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 250, Location = new Point(115, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);

                pName.Controls.Add(new Label { Text = "單位：", AutoSize = true, Location = new Point(385, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                TextBox txtUnit = new TextBox { Width = 100, Location = new Point(465, 7), Font = new Font("Microsoft JhengHei UI", 12F), Text = "mg/L" };
                pName.Controls.Add(txtUnit);
                
                flpEditor.Controls.Add(pName);

                FlowLayoutPanel flpSourcesContainer = new FlowLayoutPanel { Width = 950, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                flpEditor.Controls.Add(flpSourcesContainer);

                // 複製一份記憶體來編輯
                var editingConfigs = new List<TestConfigItem>();
                foreach (var c in _configs) {
                    editingConfigs.Add(c.Clone()); 
                }

                TestConfigItem currentEditingItem = null;
                Dictionary<string, List<string>> _columnCache = new Dictionary<string, List<string>>();
                bool isSyncing = false; 

                Action renderRows = null;
                renderRows = () => {
                    flpSourcesContainer.SuspendLayout();
                    flpSourcesContainer.Controls.Clear();

                    if (currentEditingItem == null) {
                        Label lblEmpty = new Label { Text = "請從左側點擊「➕ 新增左側統計指標」或選擇現有項目進行編輯。", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(20) };
                        flpSourcesContainer.Controls.Add(lblEmpty);
                        flpSourcesContainer.ResumeLayout(true);
                        return;
                    }

                    // 畫標題列 (寬度縮減為 920，防止水平捲軸)
                    Panel pnlHeader = new Panel { Width = 920, Height = 30 };
                    string[] headers = { "刪除", "來源資料庫", "來源資料表", "參照資料欄", "選項過濾條件", "計算欄位", "運算方式" };
                    int[] xs = { 0, 40, 150, 300, 450, 600, 750 };
                    int[] ws = { 35, 100, 140, 140, 140, 140, 100 };
                    
                    for (int i = 0; i < headers.Length; i++) {
                        Label lbl = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Location = new Point(xs[i], 5), AutoSize = true };
                        pnlHeader.Controls.Add(lbl);
                    }
                    flpSourcesContainer.Controls.Add(pnlHeader);

                    // 畫資料列
                    for (int i = 0; i < currentEditingItem.Sources.Count; i++) {
                        int currentIndex = i;
                        var srcDef = currentEditingItem.Sources[i];

                        Panel pRow = new Panel { Width = 920, Height = 50, BackColor = Color.FromArgb(245, 250, 245), Margin = new Padding(0, 5, 0, 5) };
                        pRow.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pRow.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                        
                        int cy = 8;

                        Button btnRemove = new Button { Text = "❌", Location = new Point(xs[0], cy-2), Width = ws[0], Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        btnRemove.FlatAppearance.BorderSize = 0;
                        btnRemove.Click += (s, ev) => { currentEditingItem.Sources.RemoveAt(currentIndex); renderRows(); };

                        ComboBox cbDb = new ComboBox { Location = new Point(xs[1], cy), Width = ws[1], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbTb = new ComboBox { Location = new Point(xs[2], cy), Width = ws[2], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        
                        ComboBox cbRefCol = new ComboBox { Location = new Point(xs[3], cy), Width = ws[3], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbFilter = new ComboBox { Location = new Point(xs[4], cy), Width = ws[4], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbCol = new ComboBox { Location = new Point(xs[5], cy), Width = ws[5], DropDownStyle = ComboBoxStyle.DropDown, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbAgg = new ComboBox { Location = new Point(xs[6], cy), Width = ws[6]+50, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };

                        pRow.Controls.AddRange(new Control[] { btnRemove, cbDb, cbTb, cbRefCol, cbFilter, cbCol, cbAgg });

                        cbDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                        foreach (var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

                        bool isInitializing = true;
                        bool colsLoaded = false;

                        Action lazyLoadCols = () => {
                            if (colsLoaded) return;
                            cbCol.Items.Clear(); cbCol.Items.Add("Id (無條件計數)");
                            cbRefCol.Items.Clear(); cbRefCol.Items.Add(""); 

                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName)) {
                                var cols = DataManager.GetColumnNames(selDb.EnName, selTb.EnName).Where(c => c != "Id");
                                foreach (var c in cols) {
                                    cbCol.Items.Add(c);
                                    cbRefCol.Items.Add(c); 
                                }
                            }
                            colsLoaded = true;
                        };

                        EventHandler triggerLoad = (s, ev) => { lazyLoadCols(); };
                        cbCol.DropDown += triggerLoad;
                        cbRefCol.DropDown += triggerLoad;

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
                                cbCol.Text = "Id (無條件計數)"; cbRefCol.Text = ""; cbFilter.Text = "非空值 (有輸入即算)";
                            }
                        };

                        cbRefCol.SelectedIndexChanged += (s, ev) => {
                            cbFilter.Items.Clear(); 
                            cbFilter.Items.Add("非空值 (有輸入即算)");
                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;
                            string refColStr = cbRefCol.Text;
                            
                            if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName) && !string.IsNullOrEmpty(refColStr)) {
                                string tbName = selTb.EnName;
                                string dbName = selDb.EnName;
                                
                                string multiKey = $"{tbName}|{refColStr}";
                                if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey)) {
                                    foreach (var opt in App_DropdownManager.MultiSelectCache[multiKey]) {
                                        if (!string.IsNullOrWhiteSpace(opt.Text) && !cbFilter.Items.Contains(opt.Text.Trim())) {
                                            cbFilter.Items.Add(opt.Text.Trim());
                                        }
                                    }
                                }
                                
                                string[] dropOpts = App_DropdownManager.GetAllOptionsForColumn(tbName, refColStr);
                                if (dropOpts != null && dropOpts.Length > 0) {
                                    foreach (var opt in dropOpts) {
                                        if (!string.IsNullOrWhiteSpace(opt) && !cbFilter.Items.Contains(opt.Trim())) {
                                            cbFilter.Items.Add(opt.Trim());
                                        }
                                    }
                                }

                                try {
                                    DataTable dtDist = DataManager.GetTableData(dbName, tbName, "", "", "");
                                    if (dtDist != null && dtDist.Columns.Contains(refColStr)) {
                                        var distinctVals = dtDist.Rows.Cast<DataRow>()
                                                                 .Select(r => r[refColStr]?.ToString().Trim())
                                                                 .Where(str => !string.IsNullOrEmpty(str))
                                                                 .Distinct();
                                        foreach (var val in distinctVals) {
                                            if (!cbFilter.Items.Contains(val)) {
                                                cbFilter.Items.Add(val);
                                            }
                                        }
                                    }
                                } catch { }
                            }
                            cbFilter.SelectedIndex = 0;
                        };

                        cbAgg.Items.AddRange(new string[] { "計數", "平均值", "最大值", "最小值" });
                        cbAgg.SelectedIndexChanged += (s, ev) => {
                            if (!isInitializing) {
                                if (cbAgg.Text == "平均值") srcDef.AggType = "AVG";
                                else if (cbAgg.Text == "最大值") srcDef.AggType = "MAX";
                                else if (cbAgg.Text == "最小值") srcDef.AggType = "MIN";
                                else srcDef.AggType = "COUNT";
                            }
                        };

                        // 🟢 即時同步至記憶體
                        cbCol.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.ColName = cbCol.Text; } };
                        cbFilter.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.FilterValue = cbFilter.Text; } };
                        cbRefCol.TextChanged += (s, ev) => { if (!isInitializing) { srcDef.RefColName = cbRefCol.Text; } };

                        // 初始化填值 (靜默進行)
                        foreach (ItemMap im in cbDb.Items) if (im.EnName == srcDef.DbName) { cbDb.SelectedItem = im; break; }
                        if (cbDb.SelectedItem != null && _dbMap.ContainsKey(srcDef.DbName)) {
                            cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                            foreach (var tb in _dbMap[srcDef.DbName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            foreach (ItemMap im in cbTb.Items) if (im.EnName == srcDef.TableName) { cbTb.SelectedItem = im; break; }
                        }

                        cbCol.Text = srcDef.ColName;
                        if (!string.IsNullOrEmpty(srcDef.RefColName)) cbRefCol.Text = srcDef.RefColName;

                        if (!string.IsNullOrEmpty(srcDef.FilterValue)) {
                            if (!cbFilter.Items.Contains(srcDef.FilterValue)) cbFilter.Items.Add(srcDef.FilterValue);
                            cbFilter.Text = srcDef.FilterValue;
                        } else {
                            cbFilter.Text = "非空值 (有輸入即算)";
                        }
                        
                        if (srcDef.AggType == "COUNT") cbAgg.Text = "計數";
                        else if (srcDef.AggType == "AVG") cbAgg.Text = "平均值";
                        else if (srcDef.AggType == "MAX") cbAgg.Text = "最大值";
                        else if (srcDef.AggType == "MIN") cbAgg.Text = "最小值";
                        else cbAgg.Text = "平均值"; 

                        isInitializing = false;
                        flpSourcesContainer.Controls.Add(pRow);
                    }

                    Button btnAdd = new Button { Text = "➕ 新增一行【資料對應規則】", Width = 920, Height = 45, Margin = new Padding(0, 10, 0, 0), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    btnAdd.Click += (s, ev) => { 
                        currentEditingItem.Sources.Add(new DataSourceDef()); 
                        renderRows(); 
                    };
                    
                    flpSourcesContainer.Controls.Add(btnAdd);
                    flpSourcesContainer.ResumeLayout(true);
                };

                // 🟢 綁定文字框同步機制 (即時反映至左側清單與記憶體模型)
                txtName.TextChanged += (s, ev) => {
                    if (isSyncing || currentEditingItem == null || lbItems.SelectedIndex < 0) return;
                    currentEditingItem.DisplayName = txtName.Text;
                    int idx = lbItems.SelectedIndex;
                    isSyncing = true;
                    lbItems.Items[idx] = string.IsNullOrWhiteSpace(txtName.Text) ? "[未命名]" : txtName.Text;
                    isSyncing = false;
                };

                txtUnit.TextChanged += (s, ev) => {
                    if (isSyncing || currentEditingItem == null) return;
                    currentEditingItem.Unit = txtUnit.Text;
                };

                Action bindRightPanel = () => {
                    if (currentEditingItem == null) {
                        txtName.Text = ""; txtUnit.Text = "";
                        txtName.Enabled = false; txtUnit.Enabled = false;
                    } else {
                        txtName.Enabled = true; txtUnit.Enabled = true;
                        isSyncing = true;
                        txtName.Text = currentEditingItem.DisplayName;
                        txtUnit.Text = string.IsNullOrEmpty(currentEditingItem.Unit) ? "mg/L" : currentEditingItem.Unit;
                        isSyncing = false;
                    }
                    renderRows();
                };

                Action refreshList = () => {
                    isSyncing = true;
                    lbItems.Items.Clear();
                    foreach (var cfg in editingConfigs) {
                        if (cfg != null) {
                            lbItems.Items.Add(string.IsNullOrWhiteSpace(cfg.DisplayName) ? "[未命名]" : cfg.DisplayName);
                        }
                    }
                    isSyncing = false;
                };

                btnAddNewStat.Click += (s, ev) => {
                    var newItem = new TestConfigItem();
                    newItem.Sources.Add(new DataSourceDef());
                    editingConfigs.Add(newItem);
                    refreshList();
                    lbItems.SelectedIndex = editingConfigs.Count - 1;
                    txtName.Focus();
                };

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (isSyncing || lbItems.SelectedIndex < 0) return;
                    currentEditingItem = editingConfigs[lbItems.SelectedIndex];
                    bindRightPanel();
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        editingConfigs.RemoveAt(lbItems.SelectedIndex);
                        currentEditingItem = null;
                        refreshList();
                        bindRightPanel();
                    }
                };

                // 🟢 徹底簡化儲存流程：只保留最下方一顆「儲存設定並關閉視窗」按鈕
                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);
                f.Controls.Add(tlp);

                Button btnSaveAll = new Button { Text = "💾 儲存所有設定並關閉", Dock = DockStyle.Bottom, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSaveAll.FlatAppearance.BorderSize = 0;
                
                btnSaveAll.Click += async (senderObj, evnt) => {
                    _configs = new List<TestConfigItem>(editingConfigs);
                    SaveSettings();
                    await LoadDashboardDataAsync(); 
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(btnSaveAll);
                
                refreshList();
                bindRightPanel();

                f.ShowDialog();
            }
        }

        // =========================================================
        // PDF 導出 
        // =========================================================
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
            
            PdfHelper.ExportDashboardToPdf(bitmaps, "檢測數據統計表", dateStr, "檢測數據統計表");
        }

        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 450, Height = 500, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));

                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Fill, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), AutoSize = true };
                tlp.Controls.Add(lbl, 0, 0);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(15, 5, 15, 5), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
                
                clb.Items.Add("【總計】區間統計總計 (四大區塊)", true); 
                
                foreach (var kvp in _monthlyPanels) {
                    clb.Items.Add($"近三年逐月：{kvp.Key}", true);
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
                    foreach (var item in clb.CheckedItems) {
                        string text = item.ToString();
                        if (text.Contains("【總計】區間統計總計")) {
                            selectedPanels.Add(_pnlTopBox);
                        } else {
                            string key = text.Replace("近三年逐月：", "");
                            if (_monthlyPanels.ContainsKey(key)) selectedPanels.Add(_monthlyPanels[key]);
                        }
                    }
                }
            }
            return selectedPanels;
        }
    }
}
