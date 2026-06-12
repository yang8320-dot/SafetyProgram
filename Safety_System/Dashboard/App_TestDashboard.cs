/// FILE: Safety_System/Dashboard/App_TestDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
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

        // 設定檔路徑與快取 (檢測數據專屬)
        private readonly string SettingsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestDashboardSettings.txt");
        private List<TestConfigItem> _configs = new List<TestConfigItem>();

        // 查詢按鈕，用於防呆禁用
        private Button _btnSearch;

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
            public string Unit { get; set; } = "mg/L"; // 檢測數據預設單位
            public List<DataSourceDef> Sources { get; set; } = new List<DataSourceDef>();
        }

        private class DataSourceDef
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string FilterValue { get; set; } 
            public string AggType { get; set; } 
            public string ColName2 { get; set; } = ""; // 支援日期相減的起日欄位
            public string RefColName { get; set; } = ""; // 參照資料欄 (用於過濾條件)
        }

        public Control GetView()
        {
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

        // =========================================================
        // 萬能日期解析器：自動應付民國年、西元年、各種格式差異
        // =========================================================
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

        // =========================================================
        // 日期控制與資料運算
        // =========================================================
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
                                else if (src.AggType == "SUM" || src.AggType == "AVG" || src.AggType == "MAX" || src.AggType == "MIN")
                                {
                                    if (src.ColName != "Id (無條件計數)" && dt.Columns.Contains(src.ColName)) {
                                        string valStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                        if (double.TryParse(valStr.Replace(",", ""), out double v)) gatheredValues.Add(v);
                                    }
                                }
                                else if (src.AggType.StartsWith("DIFF"))
                                {
                                    if (dt.Columns.Contains(src.ColName) && dt.Columns.Contains(src.ColName2)) {
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
                    }
                    catch { } 
                }

                double finalVal = 0;
                if (primaryAgg == "COUNT") finalVal = countVal;
                else if (primaryAgg == "SUM" || primaryAgg == "DIFF_SUM") finalVal = gatheredValues.Sum();
                else if (primaryAgg == "AVG" || primaryAgg == "DIFF_AVG") finalVal = gatheredValues.Any() ? Math.Round(gatheredValues.Average(), 2) : 0;
                else if (primaryAgg == "MAX") finalVal = gatheredValues.Any() ? gatheredValues.Max() : 0;
                else if (primaryAgg == "MIN") finalVal = gatheredValues.Any() ? gatheredValues.Min() : 0;

                result[cfg.DisplayName] = finalVal;
            }
            return result;
        }

        // =========================================================
        // 設定檔管理與動態設定視窗
        // =========================================================
        private void LoadSettings()
        {
            _configs.Clear();
            if (File.Exists(SettingsFile))
            {
                try
                {
                    foreach (var line in File.ReadAllLines(SettingsFile, Encoding.UTF8))
                    {
                        var parts = line.Split('|');
                        if (parts.Length > 1)
                        {
                            string dispName = parts[0];
                            string unit = "mg/L"; 

                            if (parts.Length > 1 && !parts[1].Contains(";")) {
                                unit = parts[1];
                            }

                            TestConfigItem cfg = new TestConfigItem { DisplayName = dispName, Unit = unit };
                            
                            int srcStartIdx = (!parts[1].Contains(";")) ? 2 : 1;

                            for (int i = srcStartIdx; i < parts.Length; i++)
                            {
                                var srcParts = parts[i].Split(';');
                                if (srcParts.Length >= 4)
                                {
                                    string filter = srcParts.Length > 4 ? srcParts[4] : "";
                                    string col2 = srcParts.Length > 5 ? srcParts[5] : ""; 
                                    // 支援舊版相容讀取
                                    string refCol = srcParts.Length > 6 ? srcParts[6] : "";
                                    cfg.Sources.Add(new DataSourceDef { DbName = srcParts[0], TableName = srcParts[1], ColName = srcParts[2], AggType = srcParts[3], FilterValue = filter, ColName2 = col2, RefColName = refCol });
                                }
                            }
                            _configs.Add(cfg);
                        }
                    }
                }
                catch { }
            }
        }

        private void SaveSettings()
        {
            try
            {
                List<string> lines = new List<string>();
                foreach (var cfg in _configs)
                {
                    if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                    string line = $"{cfg.DisplayName}|{cfg.Unit}";
                    foreach (var src in cfg.Sources)
                    {
                        // 將 RefColName 加入儲存格式
                        line += $"|{src.DbName};{src.TableName};{src.ColName};{src.AggType};{src.FilterValue};{src.ColName2};{src.RefColName}";
                    }
                    lines.Add(line);
                }
                File.WriteAllLines(SettingsFile, lines, Encoding.UTF8);
            }
            catch { }
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
                Label l2 = new Label { Text = "編輯 / 新增項目", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.SaddleBrown, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                Panel pName = new Panel { Width = 1100, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 250, Location = new Point(115, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);

                pName.Controls.Add(new Label { Text = "單位：", AutoSize = true, Location = new Point(385, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                TextBox txtUnit = new TextBox { Width = 100, Location = new Point(465, 7), Font = new Font("Microsoft JhengHei UI", 12F), Text = "mg/L" }; // 檢測數據預設為 mg/L
                pName.Controls.Add(txtUnit);
                
                Button btnAddSource = new Button { Text = "➕ 新增項目", Location = new Point(595, 5), Size = new Size(130, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnAddSource.FlatAppearance.BorderSize = 0;
                pName.Controls.Add(btnAddSource);
                
                flpEditor.Controls.Add(pName);

                FlowLayoutPanel flpSourcesContainer = new FlowLayoutPanel { Width = 1150, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                flpEditor.Controls.Add(flpSourcesContainer);

                var dbMap = App_DbConfig.GetDbMapCache();

                Action<DataSourceDef> addSourceRow = (def) => {
                    // 面板加寬以容納新欄位
                    Panel pRow = new Panel { Width = 1100, Height = 75, BackColor = Color.FromArgb(245, 250, 245), Margin = new Padding(0, 5, 0, 5) };
                    pRow.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pRow.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                    
                    int ly = 10;
                    int cy = 35;

                    // 重新排列欄位坐標
                    int x0 = 10, w0 = 35;   // 刪除按鈕
                    int x1 = 50, w1 = 110;  // 資料庫
                    int x2 = 165, w2 = 130; // 資料表
                    int x3 = 300, w3 = 130; // 參照資料欄
                    int x4 = 435, w4 = 150; // 選項過濾條件
                    int x5 = 590, w5 = 130; // 計算欄位
                    int x6 = 725, w6 = 100; // 運算方式
                    int x7 = 830, w7 = 110; // 主欄位迄
                    int x8 = 945, w8 = 110; // 次欄位起

                    Button btnRemove = new Button { Text = "❌", Location = new Point(x0, 34), Width = w0, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                    btnRemove.FlatAppearance.BorderSize = 0;

                    ComboBox cbDb = new ComboBox { Location = new Point(x1, cy), Width = w1, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbTb = new ComboBox { Location = new Point(x2, cy), Width = w2, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    
                    ComboBox cbRefCol = new ComboBox { Location = new Point(x3, cy), Width = w3, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbFilter = new ComboBox { Location = new Point(x4, cy), Width = w4, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbCol = new ComboBox { Location = new Point(x5, cy), Width = w5, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbAgg = new ComboBox { Location = new Point(x6, cy), Width = w6, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    
                    Label lblColEnd = new Label { Text = "主欄位(迄)", Location = new Point(x7, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) };
                    ComboBox cbColEnd = new ComboBox { Location = new Point(x7, cy), Width = w7, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Visible = false };
                    
                    Label lblColStart = new Label { Text = "次欄位(起)", Location = new Point(x8, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), ForeColor = Color.DarkOrange, Visible = false };
                    ComboBox cbColStart = new ComboBox { Location = new Point(x8, cy), Width = w8, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Visible = false };

                    pRow.Controls.AddRange(new Control[] {
                        btnRemove,
                        new Label { Text = "資料庫", Location = new Point(x1, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbDb,
                        new Label { Text = "資料表", Location = new Point(x2, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbTb,
                        new Label { Text = "參照資料欄", Location = new Point(x3, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbRefCol,
                        new Label { Text = "選項過濾條件", Location = new Point(x4, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbFilter,
                        new Label { Text = "計算欄位", Location = new Point(x5, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbCol,
                        new Label { Text = "運算方式", Location = new Point(x6, ly), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbAgg,
                        lblColEnd, cbColEnd,
                        lblColStart, cbColStart
                    });

                    btnRemove.Click += (s, ev) => flpSourcesContainer.Controls.Remove(pRow);

                    cbDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                    foreach (var kvp in dbMap) cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

                    cbDb.SelectedIndexChanged += (s, ev) => {
                        cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                        var selDb = cbDb.SelectedItem as ItemMap;
                        if (selDb != null && !string.IsNullOrEmpty(selDb.EnName) && dbMap.ContainsKey(selDb.EnName)) {
                            foreach (var tb in dbMap[selDb.EnName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                        }
                    };

                    cbTb.SelectedIndexChanged += (s, ev) => {
                        cbCol.Items.Clear(); cbCol.Items.Add("Id (無條件計數)");
                        cbRefCol.Items.Clear(); cbRefCol.Items.Add(""); 
                        cbColEnd.Items.Clear(); cbColEnd.Items.Add("");
                        cbColStart.Items.Clear(); cbColStart.Items.Add("");
                        
                        var selDb = cbDb.SelectedItem as ItemMap;
                        var selTb = cbTb.SelectedItem as ItemMap;
                        if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName)) {
                            var cols = DataManager.GetColumnNames(selDb.EnName, selTb.EnName).Where(c => c != "Id");
                            foreach (var c in cols) {
                                cbCol.Items.Add(c);
                                cbRefCol.Items.Add(c); 
                                cbColEnd.Items.Add(c);
                                cbColStart.Items.Add(c);
                            }
                        }
                    };

                    cbAgg.SelectedIndexChanged += (s, ev) => {
                        bool isDiff = cbAgg.Text.Contains("相減");
                        lblColEnd.Visible = isDiff;
                        cbColEnd.Visible = isDiff;
                        lblColStart.Visible = isDiff;
                        cbColStart.Visible = isDiff;
                    };

                    // 🟢 連動改由 RefCol 驅動，且動態從資料庫撈取唯一值
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

                            // 🟢 核心修復：主動從資料庫撈取不重複的值填入，以防該欄位未設定在下拉選單中
                            try {
                                DataTable dtDist = DataManager.GetTableData(dbName, tbName, "", "", "");
                                if (dtDist != null && dtDist.Columns.Contains(refColStr)) {
                                    var distinctVals = dtDist.AsEnumerable()
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

                    cbAgg.Items.AddRange(new string[] { "計數", "加總", "平均值", "最大值", "最小值", "日期相減(總天數)" });

                    if (def != null) {
                        foreach(ItemMap im in cbDb.Items) if(im.EnName == def.DbName) { cbDb.SelectedItem = im; break; }
                        foreach(ItemMap im in cbTb.Items) if(im.EnName == def.TableName) { cbTb.SelectedItem = im; break; }
                        
                        cbCol.Text = def.ColName;
                        if (!string.IsNullOrEmpty(def.RefColName) && cbRefCol.Items.Contains(def.RefColName)) {
                            cbRefCol.SelectedItem = def.RefColName;
                        }

                        if (!string.IsNullOrEmpty(def.ColName2)) {
                            cbColEnd.Text = def.ColName;
                            cbColStart.Text = def.ColName2;
                        }

                        // 🟢 確保 FilterValue 被正確載入
                        if (!string.IsNullOrEmpty(def.FilterValue)) {
                            if (!cbFilter.Items.Contains(def.FilterValue)) cbFilter.Items.Add(def.FilterValue);
                            cbFilter.Text = def.FilterValue;
                        } else {
                            cbFilter.Text = "非空值 (有輸入即算)";
                        }
                        
                        if (def.AggType == "COUNT") cbAgg.Text = "計數";
                        else if (def.AggType == "SUM") cbAgg.Text = "加總";
                        else if (def.AggType == "AVG") cbAgg.Text = "平均值";
                        else if (def.AggType == "MAX") cbAgg.Text = "最大值";
                        else if (def.AggType == "MIN") cbAgg.Text = "最小值";
                        else if (def.AggType == "DIFF_SUM") cbAgg.Text = "日期相減(總天數)";
                        else cbAgg.Text = "平均值"; 
                    } else {
                        cbAgg.Text = "平均值";
                        cbFilter.Text = "非空值 (有輸入即算)";
                    }

                    flpSourcesContainer.Controls.Add(pRow);
                };

                btnAddSource.Click += (s, ev) => addSourceRow(null);

                Button btnSaveRow = new Button { Text = "💾 儲存並加入清單", Width = 900, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 0), Cursor = Cursors.Hand };

                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);
                pnlRight.Controls.Add(btnSaveRow);
                btnSaveRow.Dock = DockStyle.Bottom;

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);
                f.Controls.Add(tlp);

                // 資料載入與綁定邏輯
                Action refreshList = () => {
                    lbItems.Items.Clear();
                    foreach (var cfg in _configs) {
                        if (cfg != null && !string.IsNullOrEmpty(cfg.DisplayName)) {
                            lbItems.Items.Add(cfg.DisplayName);
                        }
                    }
                };
                refreshList();

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (lbItems.SelectedIndex < 0) return;
                    flpSourcesContainer.Controls.Clear();
                    var cfg = _configs[lbItems.SelectedIndex];
                    txtName.Text = cfg.DisplayName;
                    txtUnit.Text = string.IsNullOrEmpty(cfg.Unit) ? "mg/L" : cfg.Unit; 
                    
                    foreach (var src in cfg.Sources) {
                        addSourceRow(src);
                    }
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        _configs.RemoveAt(lbItems.SelectedIndex);
                        SaveSettings();
                        refreshList();
                        txtName.Clear();
                        flpSourcesContainer.Controls.Clear();
                    }
                };

                btnSaveRow.Click += (ss, ee) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入顯示名稱！"); return; }
                    
                    var newCfg = new TestConfigItem { 
                        DisplayName = txtName.Text.Trim(),
                        Unit = string.IsNullOrWhiteSpace(txtUnit.Text) ? "mg/L" : txtUnit.Text.Trim()
                    };
                    
                    foreach (Control ctrl in flpSourcesContainer.Controls) {
                        if (ctrl is Panel pRow) {
                            var cbDb = pRow.Controls[2] as ComboBox;
                            var cbTb = pRow.Controls[4] as ComboBox;
                            var cbRefCol = pRow.Controls[6] as ComboBox; 
                            var cbFilter = pRow.Controls[8] as ComboBox;
                            var cbCol = pRow.Controls[10] as ComboBox; 
                            var cbAgg = pRow.Controls[12] as ComboBox;
                            var cbColEnd = pRow.Controls[14] as ComboBox;
                            var cbColStart = pRow.Controls[16] as ComboBox;

                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;

                            if (selDb != null && selTb != null && !string.IsNullOrWhiteSpace(cbCol.Text) && !string.IsNullOrWhiteSpace(cbAgg.Text)) {
                                string filterVal = (cbFilter.Text == "非空值 (有輸入即算)") ? "" : cbFilter.Text;
                                
                                string aggTypeDb = "COUNT";
                                string dbColToSave = cbCol.Text;
                                string dbCol2ToSave = "";

                                if (cbAgg.Text == "加總") {
                                    aggTypeDb = "SUM";
                                }
                                else if (cbAgg.Text == "平均值") {
                                    aggTypeDb = "AVG";
                                }
                                else if (cbAgg.Text == "最大值") {
                                    aggTypeDb = "MAX";
                                }
                                else if (cbAgg.Text == "最小值") {
                                    aggTypeDb = "MIN";
                                }
                                else if (cbAgg.Text == "日期相減(總天數)") {
                                    aggTypeDb = "DIFF_SUM";
                                    dbColToSave = cbColEnd.Text;
                                    dbCol2ToSave = cbColStart.Text;
                                }
                                
                                newCfg.Sources.Add(new DataSourceDef { 
                                    DbName = selDb.EnName, 
                                    TableName = selTb.EnName, 
                                    ColName = dbColToSave, 
                                    FilterValue = filterVal, 
                                    AggType = aggTypeDb,
                                    ColName2 = dbCol2ToSave,
                                    RefColName = cbRefCol.Text
                                });
                            }
                        }
                    }

                    if (newCfg.Sources.Count == 0) { MessageBox.Show("請至少設定一組完整的資料來源！"); return; }

                    int existingIdx = _configs.FindIndex(c => c != null && c.DisplayName == newCfg.DisplayName);
                    if (existingIdx >= 0) _configs[existingIdx] = newCfg;
                    else _configs.Add(newCfg);

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
    }
}
