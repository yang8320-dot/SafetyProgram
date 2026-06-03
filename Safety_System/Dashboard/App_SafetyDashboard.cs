/// FILE: Safety_System/Dashboard/App_SafetyDashboard.cs ///
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
    public class App_SafetyDashboard
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

        // 設定檔路徑與快取
        private readonly string SettingsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetyDashboardSettings.txt");
        private List<SafetyConfigItem> _configs = new List<SafetyConfigItem>();

        // 查詢按鈕，用於防呆禁用
        private Button _btnSearch;

        // 定義下拉選單對應的中英文模型
        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        // 定義自訂統計項目的資料結構 (動態無限來源，並加入「單位」)
        private class SafetyConfigItem
        {
            public string DisplayName { get; set; }
            public string Unit { get; set; } = "件"; // 預設單位
            public List<DataSourceDef> Sources { get; set; } = new List<DataSourceDef>();
        }

        private class DataSourceDef
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string FilterValue { get; set; } 
            public string AggType { get; set; } 
        }

        public Control GetView()
        {
            LoadSettings();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            
            // 使用嚴格的 TableLayoutPanel 確保由上到下絕對不會亂掉
            TableLayoutPanel masterLayout = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1, 
                RowCount = 4,
                Margin = new Padding(0)
            };
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 標題
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 條件按鈕
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 四大區塊
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 逐月矩陣表

            // ==========================================
            // 第一行：大標題
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "🛡️ 工安事件與指標綜合數據看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.SteelBlue, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
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

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(100, btnHeight), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(130, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
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

            Color headColor = Color.SteelBlue;
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
            // 防重複點擊，導致進程交錯崩潰
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
                                
                                var mResult = CalculatePeriodStats(mStart, mEnd, new List<SafetyConfigItem> { cfg });
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

                // ================= UI 更新 (上半部區塊) =================
                _pnlData1.Controls.Clear(); _pnlData2.Controls.Clear(); _pnlData3.Controls.Clear(); _pnlData4.Controls.Clear();
                foreach (var cfg in _configs)
                {
                    if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;

                    string key = cfg.DisplayName;
                    string unit = string.IsNullOrEmpty(cfg.Unit) ? "件" : cfg.Unit;

                    double vCurr = dictCurr.ContainsKey(key) ? dictCurr[key] : 0;
                    double vLy = dictLy.ContainsKey(key) ? dictLy[key] : 0;
                    double vL2y = dictL2y.ContainsKey(key) ? dictL2y[key] : 0;

                    _pnlData1.Controls.Add(CreateStatLabel(key, vCurr, unit));
                    _pnlData2.Controls.Add(CreateStatLabel(key, vLy, unit));
                    _pnlData3.Controls.Add(CreateStatLabel(key, vL2y, unit));

                    double diff = vCurr - vLy;
                    string diffText = (diff > 0 ? "+" : "") + diff.ToString("N0") + " " + unit;
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

                        int targetWidth = 1000; 
                        if (_flpBottomBox.Width > 40) targetWidth = _flpBottomBox.Width - 20;
                        else if (_flpBottomBox.Parent != null && _flpBottomBox.Parent.Width > 40) targetWidth = _flpBottomBox.Parent.Width - 40;

                        Panel pnlWrapper = new Panel { Width = targetWidth, Height = 220, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
                        pnlWrapper.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlWrapper.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                        
                        Label lblTitle = new Label { Text = $"📊 近三年逐月統計：{statName}", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Padding = new Padding(15, 10, 0, 10), Dock = DockStyle.Top };
                        
                        DataGridView dgv = new DataGridView { 
                            Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                            BorderStyle = BorderStyle.None, Margin = new Padding(10)
                        };
                        
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkSlateBlue;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dgv.ColumnHeadersHeight = 35;
                        dgv.RowTemplate.Height = 33;
                        dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                        
                        dgv.DataSource = dt;

                        if (dgv.Columns.Contains("年度")) {
                            dgv.Columns["年度"].Width = 100;
                        }
                        
                        if (dgv.Columns.Contains("年度總計")) {
                            dgv.Columns["年度總計"].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                            dgv.Columns["年度總計"].DefaultCellStyle.BackColor = Color.LightYellow;
                            dgv.Columns["年度總計"].DefaultCellStyle.Format = "N0";
                        }
                        
                        for (int i = 1; i <= 12; i++) {
                            string monthCol = $"{i}月";
                            if (dgv.Columns.Contains(monthCol)) {
                                dgv.Columns[monthCol].DefaultCellStyle.Format = "N0";
                            }
                        }
                        
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

        private Label CreateStatLabel(string title, double value, string unit)
        {
            return new Label { Text = $"{title}: {value:N0} {unit}", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        private Dictionary<string, double> CalculatePeriodStats(DateTime sDate, DateTime eDate, List<SafetyConfigItem> targetConfigs = null)
        {
            var result = new Dictionary<string, double>();
            var configsToRun = targetConfigs ?? _configs;
            string sStr = sDate.ToString("yyyy-MM-dd");
            string eStr = eDate.ToString("yyyy-MM-dd");

            foreach (var cfg in configsToRun)
            {
                if (cfg == null || string.IsNullOrEmpty(cfg.DisplayName)) continue;
                
                double totalVal = 0;
                if (cfg.Sources == null) continue; 

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

                        foreach (DataRow r in dt.Rows)
                        {
                            if (r.RowState == DataRowState.Deleted) continue;
                            
                            bool match = false;
                            string valStr = "";

                            if (!string.IsNullOrEmpty(src.ColName) && src.ColName != "Id (無條件計數)" && dt.Columns.Contains(src.ColName))
                            {
                                valStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(filterTarget)) {
                                    match = !string.IsNullOrEmpty(valStr);
                                } else {
                                    // 🟢 強化精準度：去頭尾空白，並且忽略大小寫比對，防止資料庫內有些微打字差異
                                    match = valStr.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries)
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
                                    totalVal++;
                                }
                                else if (src.AggType == "SUM")
                                {
                                    if (double.TryParse(valStr.Replace(",", ""), out double v)) totalVal += v;
                                }
                            }
                        }
                    }
                    catch { } 
                }
                result[cfg.DisplayName] = totalVal;
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
                            string unit = "件"; 

                            if (parts.Length > 1 && !parts[1].Contains(";")) {
                                unit = parts[1];
                            }

                            SafetyConfigItem cfg = new SafetyConfigItem { DisplayName = dispName, Unit = unit };
                            
                            int srcStartIdx = (!parts[1].Contains(";")) ? 2 : 1;

                            for (int i = srcStartIdx; i < parts.Length; i++)
                            {
                                var srcParts = parts[i].Split(';');
                                if (srcParts.Length >= 4)
                                {
                                    string filter = srcParts.Length > 4 ? srcParts[4] : "";
                                    cfg.Sources.Add(new DataSourceDef { DbName = srcParts[0], TableName = srcParts[1], ColName = srcParts[2], AggType = srcParts[3], FilterValue = filter });
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
                        line += $"|{src.DbName};{src.TableName};{src.ColName};{src.AggType};{src.FilterValue}";
                    }
                    lines.Add(line);
                }
                File.WriteAllLines(SettingsFile, lines, Encoding.UTF8);
            }
            catch { }
        }

        private void BtnSetting_Click(object sender, EventArgs e)
        {
            // 🟢 取消密碼授權 (移除 AuthManager 驗證)

            using (Form f = new Form { Text = "⚙️ 看板自訂統計項目設定", Size = new Size(1400, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
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
                Label l2 = new Label { Text = "編輯 / 新增項目", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                // 🟢 間距修正：再加大 10px 
                Panel pName = new Panel { Width = 1000, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 230, Location = new Point(115, 7), Font = new Font("Microsoft JhengHei UI", 12F) };
                pName.Controls.Add(txtName);

                // 單位也跟著往後移，避免被蓋住 (增加間距)
                pName.Controls.Add(new Label { Text = "單位：", AutoSize = true, Location = new Point(370, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtUnit = new TextBox { Width = 100, Location = new Point(435, 7), Font = new Font("Microsoft JhengHei UI", 12F), Text = "件" }; 
                pName.Controls.Add(txtUnit);
                
                Button btnAddSource = new Button { Text = "➕ 新增項目", Location = new Point(555, 5), Size = new Size(130, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnAddSource.FlatAppearance.BorderSize = 0;
                pName.Controls.Add(btnAddSource);
                
                flpEditor.Controls.Add(pName);

                FlowLayoutPanel flpSourcesContainer = new FlowLayoutPanel { Width = 1050, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                flpEditor.Controls.Add(flpSourcesContainer);

                var dbMap = App_DbConfig.GetDbMapCache();

                Action<DataSourceDef> addSourceRow = (def) => {
                    Panel pRow = new Panel { Width = 1020, Height = 80, BackColor = Color.FromArgb(245, 250, 245), Margin = new Padding(0, 5, 0, 5) };
                    pRow.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pRow.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                    
                    ComboBox cbDb = new ComboBox { Width = 160, Location = new Point(10, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbTb = new ComboBox { Width = 200, Location = new Point(180, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbCol = new ComboBox { Width = 180, Location = new Point(390, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbFilter = new ComboBox { Width = 200, Location = new Point(580, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbAgg = new ComboBox { Width = 100, Location = new Point(790, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    Button btnRemove = new Button { Text = "❌", Width = 40, Height = 30, Location = new Point(900, 34), BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                    btnRemove.FlatAppearance.BorderSize = 0;

                    pRow.Controls.AddRange(new Control[] {
                        new Label { Text = "資料庫", Location = new Point(10, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbDb,
                        new Label { Text = "資料表", Location = new Point(180, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbTb,
                        new Label { Text = "計算欄位", Location = new Point(390, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbCol,
                        new Label { Text = "選項過濾條件", Location = new Point(580, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbFilter,
                        new Label { Text = "運算方式", Location = new Point(790, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) }, cbAgg,
                        btnRemove
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
                        var selDb = cbDb.SelectedItem as ItemMap;
                        var selTb = cbTb.SelectedItem as ItemMap;
                        if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName)) {
                            var cols = DataManager.GetColumnNames(selDb.EnName, selTb.EnName).Where(c => c != "Id");
                            foreach (var c in cols) cbCol.Items.Add(c);
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
                                    if (!string.IsNullOrWhiteSpace(opt) && !cbFilter.Items.Contains(opt.Trim())) {
                                        cbFilter.Items.Add(opt.Trim());
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

                    cbAgg.Items.AddRange(new string[] { "計數", "加總" });

                    if (def != null) {
                        foreach(ItemMap im in cbDb.Items) if(im.EnName == def.DbName) { cbDb.SelectedItem = im; break; }
                        foreach(ItemMap im in cbTb.Items) if(im.EnName == def.TableName) { cbTb.SelectedItem = im; break; }
                        
                        cbCol.Text = def.ColName;
                        if (!string.IsNullOrEmpty(def.FilterValue) && cbFilter.Items.Contains(def.FilterValue)) {
                            cbFilter.Text = def.FilterValue;
                        } else {
                            if (!string.IsNullOrEmpty(def.FilterValue)) cbFilter.Items.Add(def.FilterValue);
                            cbFilter.Text = string.IsNullOrEmpty(def.FilterValue) ? "非空值 (有輸入即算)" : def.FilterValue;
                        }
                        
                        if (def.AggType == "COUNT") cbAgg.Text = "計數";
                        else if (def.AggType == "SUM") cbAgg.Text = "加總";
                        else cbAgg.Text = "計數";
                    } else {
                        cbAgg.Text = "計數";
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
                    txtUnit.Text = string.IsNullOrEmpty(cfg.Unit) ? "件" : cfg.Unit; 
                    
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
                    
                    var newCfg = new SafetyConfigItem { 
                        DisplayName = txtName.Text.Trim(),
                        Unit = string.IsNullOrWhiteSpace(txtUnit.Text) ? "件" : txtUnit.Text.Trim()
                    };
                    
                    foreach (Control ctrl in flpSourcesContainer.Controls) {
                        if (ctrl is Panel pRow) {
                            var cbDb = pRow.Controls[1] as ComboBox;
                            var cbTb = pRow.Controls[3] as ComboBox;
                            var cbCol = pRow.Controls[5] as ComboBox;
                            var cbFilter = pRow.Controls[7] as ComboBox;
                            var cbAgg = pRow.Controls[9] as ComboBox;

                            var selDb = cbDb.SelectedItem as ItemMap;
                            var selTb = cbTb.SelectedItem as ItemMap;

                            if (selDb != null && selTb != null && !string.IsNullOrWhiteSpace(cbCol.Text) && !string.IsNullOrWhiteSpace(cbAgg.Text)) {
                                string filterVal = (cbFilter.Text == "非空值 (有輸入即算)") ? "" : cbFilter.Text;
                                string aggTypeDb = (cbAgg.Text == "加總") ? "SUM" : "COUNT";
                                
                                newCfg.Sources.Add(new DataSourceDef { DbName = selDb.EnName, TableName = selTb.EnName, ColName = cbCol.Text, FilterValue = filterVal, AggType = aggTypeDb });
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
        // PDF 導出 (🟢 優化：精準的總頁數計算，專業排版)
        // =========================================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 450, Height = 450, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Top, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                
                clb.Items.Add("【總計】區間統計總計 (四大區塊)", true); 
                
                foreach (var kvp in _monthlyPanels) {
                    clb.Items.Add($"近三年逐月：{kvp.Key}", true);
                }
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

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

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "工安數據統計表_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        Application.UseWaitCursor = true;
                        Cursor.Current = Cursors.WaitCursor;

                        List<Bitmap> bitmaps = new List<Bitmap>();
                        foreach (var pnl in panelsToExport) 
                        {
                            Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                            pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                            bitmaps.Add(bmp);
                        }

                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = true; 
                        pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                        int currentBmpIndex = 0;
                        int pageNumber = 1;

                        // 🟢 完美精準的總頁數模擬計算
                        int totalPages = 1;
                        float simW = 1169f - 80f; // A4 橫式寬度扣掉左右 Margin
                        float simH = 827f - 80f;  // A4 橫式高度扣掉上下 Margin
                        float simStartY = 40f + 145f; // Top Margin + 標頭預留高度
                        float simCurrentY = simStartY;
                        float simBottomLimit = 40f + simH - 30f; // 扣除底部頁碼空間

                        foreach (var bmp in bitmaps) {
                            float simScale = simW / bmp.Width;
                            float simScaledHeight = bmp.Height * simScale;

                            if (simCurrentY + simScaledHeight > simBottomLimit) {
                                if (simCurrentY == simStartY) {
                                    simCurrentY += simScaledHeight + 20f;
                                } else {
                                    totalPages++;
                                    simCurrentY = simStartY + simScaledHeight + 20f;
                                }
                            } else {
                                simCurrentY += simScaledHeight + 20f;
                            }
                        }

                        pd.PrintPage += (s, ev) => 
                        {
                            Graphics g = ev.Graphics;
                            float w = ev.MarginBounds.Width;
                            float x = ev.MarginBounds.Left;
                            float y = ev.MarginBounds.Top;

                            Font fTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                            Font fSub = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                            Font fSign = new Font("Microsoft JhengHei UI", 12F);
                            Font fDate = new Font("Microsoft JhengHei UI", 11F);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center };
                            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near };

                            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                            y += 40;

                            g.DrawString("工安數據統計表", fSub, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                            y += 40;

                            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    填表人：______________";
                            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                            y += 35;

                            string dateStr = $"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}        查詢區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text} ~ {_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
                            g.DrawString(dateStr, fDate, Brushes.DimGray, new RectangleF(x, y, w, 20), sfLeft); 
                            y += 30;

                            float headerHeightReserved = y; 
                            float bottomLimit = ev.MarginBounds.Bottom - 30; 

                            while (currentBmpIndex < bitmaps.Count) 
                            {
                                Bitmap bmp = bitmaps[currentBmpIndex];
                                float scale = w / bmp.Width;
                                float scaledHeight = bmp.Height * scale;

                                if (y + scaledHeight > bottomLimit) 
                                {
                                    if (y == headerHeightReserved) 
                                    {
                                        scale = Math.Min(scale, (float)(bottomLimit - y) / bmp.Height);
                                        scaledHeight = bmp.Height * scale;
                                        g.DrawImage(bmp, x, y, bmp.Width * scale, scaledHeight);
                                        y += scaledHeight + 20;
                                        currentBmpIndex++;
                                    } 
                                    else 
                                    {
                                        break; 
                                    }
                                } 
                                else 
                                {
                                    g.DrawImage(bmp, x, y, w, scaledHeight);
                                    y += scaledHeight + 20; 
                                    currentBmpIndex++;
                                }
                            }

                            g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fDate, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom - 15, w, 20), sfCenter);

                            if (currentBmpIndex < bitmaps.Count) {
                                pageNumber++;
                                ev.HasMorePages = true;
                            } else {
                                ev.HasMorePages = false;
                            }
                        };

                        pd.Print();
                        foreach (var bmp in bitmaps) bmp.Dispose();

                        // 🟢 強制在跳出訊息框前恢復游標，防止一直轉圈
                        Application.UseWaitCursor = false;
                        Cursor.Current = Cursors.Default;
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;

                        MessageBox.Show("PDF 匯出成功！已依設定格式排版完成。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        Application.UseWaitCursor = false;
                        Cursor.Current = Cursors.Default;
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } 
                }
            }
        }
    }
}
