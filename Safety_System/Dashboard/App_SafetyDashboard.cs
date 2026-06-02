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

        // 定義自訂統計項目的資料結構 (動態無限來源)
        private class SafetyConfigItem
        {
            public string DisplayName { get; set; }
            public List<DataSourceDef> Sources { get; set; } = new List<DataSourceDef>();
        }

        private class DataSourceDef
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string FilterValue { get; set; } // 🟢 新增：條件過濾 (空值、特定下拉選項)
            public string AggType { get; set; } 
        }

        public Control GetView()
        {
            LoadSettings();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            
            // 🟢 第一行：大標題
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60 };
            Label lblTitle = new Label { Text = "🛡️ 工安事件與指標綜合數據看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.SteelBlue, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblTitle);

            // 🟢 第二行：查詢及功能鍵
            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 0, 0, 20) };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes();

            Button btnSearch = new Button { Text = "🔍 查詢統計", Size = new Size(130, 32), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(150, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            Button btnSetting = new Button { Text = "⚙️ 統計項目設定", Size = new Size(160, 32), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnSetting.Click += BtnSetting_Click;

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearch, btnPdf, btnSetting
            });

            mainScrollPanel.Controls.Add(flpControls);
            mainScrollPanel.Controls.Add(pnlHeader);

            // ==========================================
            // 大框 1：區間統計與四大區塊
            // ==========================================
            _pnlTopBox = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            _pnlTopBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlTopBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel gridFour = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 4, RowCount = 2, Padding = new Padding(10) };
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
            
            // 加入佔位與排版設定
            mainScrollPanel.Controls.Add(_pnlTopBox);
            _pnlTopBox.BringToFront();
            flpControls.BringToFront();
            pnlHeader.BringToFront();

            // ==========================================
            // 大框 2：年度逐月統計 (多項獨立表格)
            // ==========================================
            _flpBottomBox = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0) };
            mainScrollPanel.Controls.Add(_flpBottomBox);
            _flpBottomBox.BringToFront();

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
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            return new DateTime((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()), day > maxDay ? maxDay : day);
        }

        private async Task LoadDashboardDataAsync()
        {
            if (_configs.Count == 0)
            {
                _pnlData1.Controls.Clear(); _pnlData2.Controls.Clear(); _pnlData3.Controls.Clear(); _pnlData4.Controls.Clear();
                _pnlData1.Controls.Add(new Label { Text = "請先點擊上方 [統計項目設定]\n新增欲追蹤的指標！", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                _flpBottomBox.Controls.Clear();
                _monthlyPanels.Clear();
                return;
            }

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try
            {
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
                    // 1. 計算四大區塊 (區間差異)
                    dictCurr = CalculatePeriodStats(dtS, dtE);
                    dictLy = CalculatePeriodStats(dtS.AddYears(-1), dtE.AddYears(-1));
                    dictL2y = CalculatePeriodStats(dtS.AddYears(-2), dtE.AddYears(-2));

                    // 2. 計算底部矩陣表 (每項一個 DataTable)
                    int baseYear = dtE.Year;
                    int[] years = { baseYear, baseYear - 1, baseYear - 2 };

                    foreach (var cfg in _configs)
                    {
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

                // 更新 UI - 四大區塊
                _pnlData1.Controls.Clear(); _pnlData2.Controls.Clear(); _pnlData3.Controls.Clear(); _pnlData4.Controls.Clear();
                foreach (var cfg in _configs)
                {
                    string key = cfg.DisplayName;
                    double vCurr = dictCurr.ContainsKey(key) ? dictCurr[key] : 0;
                    double vLy = dictLy.ContainsKey(key) ? dictLy[key] : 0;
                    double vL2y = dictL2y.ContainsKey(key) ? dictL2y[key] : 0;

                    _pnlData1.Controls.Add(CreateStatLabel(key, vCurr));
                    _pnlData2.Controls.Add(CreateStatLabel(key, vLy));
                    _pnlData3.Controls.Add(CreateStatLabel(key, vL2y));

                    double diff = vCurr - vLy;
                    string diffText = (diff > 0 ? "+" : "") + diff.ToString("N0") + " 單位";
                    Color diffColor = diff > 0 ? Color.IndianRed : (diff < 0 ? Color.ForestGreen : Color.DimGray);

                    _pnlData4.Controls.Add(new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) });
                }

                // 更新 UI - 底部 Grid (每個指標產生一個 Table)
                _flpBottomBox.Controls.Clear();
                _monthlyPanels.Clear();

                foreach (var kvp in monthlyTables)
                {
                    string statName = kvp.Key;
                    DataTable dt = kvp.Value;

                    Panel pnlWrapper = new Panel { Width = _flpBottomBox.Parent.ClientSize.Width - 40, Height = 195, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
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
                    dgv.Columns["年度"].Width = 80;
                    dgv.Columns["年度總計"].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
                    dgv.Columns["年度總計"].DefaultCellStyle.BackColor = Color.LightYellow;
                    
                    for (int i = 1; i <= 12; i++) dgv.Columns[$"{i}月"].DefaultCellStyle.Format = "N0";
                    dgv.Columns["年度總計"].DefaultCellStyle.Format = "N0";
                    dgv.ClearSelection();

                    pnlWrapper.Controls.Add(dgv);
                    pnlWrapper.Controls.Add(lblTitle);
                    
                    _flpBottomBox.Controls.Add(pnlWrapper);
                    _monthlyPanels[statName] = pnlWrapper;

                    // 確保寬度自適應
                    _flpBottomBox.Resize += (s, e) => { pnlWrapper.Width = _flpBottomBox.ClientSize.Width - 10; };
                }
            }
            finally
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private Label CreateStatLabel(string title, double value)
        {
            return new Label { Text = $"{title}: {value:N0} 單位", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        // 核心運算引擎：支援多來源聚合與「選項/非空值」過濾
        private Dictionary<string, double> CalculatePeriodStats(DateTime sDate, DateTime eDate, List<SafetyConfigItem> targetConfigs = null)
        {
            var result = new Dictionary<string, double>();
            var configsToRun = targetConfigs ?? _configs;
            string sStr = sDate.ToString("yyyy-MM-dd");
            string eStr = eDate.ToString("yyyy-MM-dd");

            foreach (var cfg in configsToRun)
            {
                double totalVal = 0;
                foreach (var src in cfg.Sources)
                {
                    if (string.IsNullOrEmpty(src.DbName) || string.IsNullOrEmpty(src.TableName)) continue;

                    try
                    {
                        var cols = DataManager.GetColumnNames(src.DbName, src.TableName);
                        string dateCol = cols.Contains("日期") ? "日期" : (cols.Contains("年月") ? "年月" : (cols.Contains("年度") ? "年度" : ""));
                        if (string.IsNullOrEmpty(dateCol)) continue;

                        DataTable dt = DataManager.GetTableData(src.DbName, src.TableName, dateCol, sStr, eStr);
                        if (dt == null) continue;

                        foreach (DataRow r in dt.Rows)
                        {
                            if (r.RowState == DataRowState.Deleted) continue;
                            
                            // 🟢 過濾邏輯
                            bool match = false;
                            string valStr = "";

                            if (!string.IsNullOrEmpty(src.ColName) && src.ColName != "Id (計數專用)" && dt.Columns.Contains(src.ColName))
                            {
                                valStr = r[src.ColName]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(src.FilterValue) || src.FilterValue == "非空值 (有輸入即算)") {
                                    match = !string.IsNullOrEmpty(valStr);
                                } else {
                                    // 判斷多選文字中是否包含該選項
                                    match = valStr.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                  .Select(x => x.Trim()).Contains(src.FilterValue);
                                }
                            }
                            else
                            {
                                // 若指定 "Id (計數專用)" 則強制算入
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
                            SafetyConfigItem cfg = new SafetyConfigItem { DisplayName = parts[0] };
                            for (int i = 1; i < parts.Length; i++)
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
                    string line = cfg.DisplayName;
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
            string authPrompt = "修改看板統計設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (Form f = new Form { Text = "⚙️ 看板自訂統計項目設定", Size = new Size(1300, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                // 左側：現有項目清單
                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                Label l1 = new Label { Text = "已建立的統計項目", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 30 };
                ListBox lbItems = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand };
                
                pnlLeft.Controls.Add(lbItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnDel);

                // 右側：編輯區
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯 / 新增項目", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
                
                Panel pName = new Panel { Width = 900, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 300, Location = new Point(100, 7), Font = new Font("Microsoft JhengHei UI", 12F) };
                pName.Controls.Add(txtName);
                
                Button btnAddSource = new Button { Text = "➕ 新增資料來源", Location = new Point(420, 5), Size = new Size(150, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnAddSource.FlatAppearance.BorderSize = 0;
                pName.Controls.Add(btnAddSource);
                
                flpEditor.Controls.Add(pName);

                FlowLayoutPanel flpSourcesContainer = new FlowLayoutPanel { Width = 950, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                flpEditor.Controls.Add(flpSourcesContainer);

                var dbMap = App_DbConfig.GetDbMapCache();

                // 🟢 動態產生來源列的方法
                Action<DataSourceDef> addSourceRow = (def) => {
                    Panel pRow = new Panel { Width = 900, Height = 80, BackColor = Color.FromArgb(245, 250, 245), Margin = new Padding(0, 5, 0, 5) };
                    pRow.Paint += (s, ev) => ControlPaint.DrawBorder(ev.Graphics, pRow.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                    
                    ComboBox cbDb = new ComboBox { Width = 140, Location = new Point(10, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbTb = new ComboBox { Width = 180, Location = new Point(160, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbCol = new ComboBox { Width = 160, Location = new Point(350, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbFilter = new ComboBox { Width = 150, Location = new Point(520, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    ComboBox cbAgg = new ComboBox { Width = 90, Location = new Point(680, 35), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    Button btnRemove = new Button { Text = "❌", Width = 40, Height = 30, Location = new Point(780, 34), BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                    btnRemove.FlatAppearance.BorderSize = 0;

                    pRow.Controls.AddRange(new Control[] {
                        new Label { Text = "資料庫", Location = new Point(10, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cbDb,
                        new Label { Text = "資料表", Location = new Point(160, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cbTb,
                        new Label { Text = "計算欄位", Location = new Point(350, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cbCol,
                        new Label { Text = "選項過濾條件", Location = new Point(520, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cbFilter,
                        new Label { Text = "運算方式", Location = new Point(680, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cbAgg,
                        btnRemove
                    });

                    btnRemove.Click += (s, ev) => flpSourcesContainer.Controls.Remove(pRow);

                    cbDb.Items.Add("");
                    foreach (var kvp in dbMap) cbDb.Items.Add(kvp.Key);

                    cbDb.SelectedIndexChanged += (s, ev) => {
                        cbTb.Items.Clear(); cbTb.Items.Add("");
                        if (dbMap.ContainsKey(cbDb.Text)) {
                            foreach (var tb in dbMap[cbDb.Text].Tables.Keys) cbTb.Items.Add(tb);
                        }
                    };

                    cbTb.SelectedIndexChanged += (s, ev) => {
                        cbCol.Items.Clear(); cbCol.Items.Add("Id (計數專用)");
                        if (dbMap.ContainsKey(cbDb.Text) && !string.IsNullOrEmpty(cbTb.Text)) {
                            var cols = DataManager.GetColumnNames(cbDb.Text, cbTb.Text).Where(c => c != "Id");
                            foreach (var c in cols) cbCol.Items.Add(c);
                        }
                    };

                    // 🟢 動態抓取下拉選單選項
                    cbCol.SelectedIndexChanged += (s, ev) => {
                        cbFilter.Items.Clear(); cbFilter.Items.Add("非空值 (有輸入即算)");
                        string tb = cbTb.Text; string col = cbCol.Text;
                        if (!string.IsNullOrEmpty(tb) && col != "Id (計數專用)") {
                            string multiKey = $"{tb}|{col}";
                            if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey)) {
                                foreach (var opt in App_DropdownManager.MultiSelectCache[multiKey]) cbFilter.Items.Add(opt);
                            } else {
                                var opts = App_DropdownManager.GetAllOptionsForColumn(tb, col);
                                if (opts != null) {
                                    foreach (var opt in opts) if (!string.IsNullOrEmpty(opt.Trim())) cbFilter.Items.Add(opt);
                                }
                            }
                        }
                        cbFilter.SelectedIndex = 0;
                    };

                    cbAgg.Items.AddRange(new string[] { "COUNT", "SUM" });

                    // 填入預設值
                    if (def != null) {
                        cbDb.Text = def.DbName;
                        cbTb.Text = def.TableName;
                        cbCol.Text = def.ColName;
                        if (!string.IsNullOrEmpty(def.FilterValue) && cbFilter.Items.Contains(def.FilterValue)) {
                            cbFilter.Text = def.FilterValue;
                        } else {
                            if (!string.IsNullOrEmpty(def.FilterValue)) cbFilter.Items.Add(def.FilterValue);
                            cbFilter.Text = string.IsNullOrEmpty(def.FilterValue) ? "非空值 (有輸入即算)" : def.FilterValue;
                        }
                        cbAgg.Text = def.AggType;
                    } else {
                        cbAgg.Text = "COUNT";
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
                    foreach (var cfg in _configs) lbItems.Items.Add(cfg.DisplayName);
                };
                refreshList();

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (lbItems.SelectedIndex < 0) return;
                    flpSourcesContainer.Controls.Clear();
                    var cfg = _configs[lbItems.SelectedIndex];
                    txtName.Text = cfg.DisplayName;
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
                    
                    var newCfg = new SafetyConfigItem { DisplayName = txtName.Text.Trim() };
                    
                    foreach (Control ctrl in flpSourcesContainer.Controls) {
                        if (ctrl is Panel pRow) {
                            var cbDb = pRow.Controls[1] as ComboBox;
                            var cbTb = pRow.Controls[3] as ComboBox;
                            var cbCol = pRow.Controls[5] as ComboBox;
                            var cbFilter = pRow.Controls[7] as ComboBox;
                            var cbAgg = pRow.Controls[9] as ComboBox;

                            if (!string.IsNullOrWhiteSpace(cbDb.Text) && !string.IsNullOrWhiteSpace(cbTb.Text) && !string.IsNullOrWhiteSpace(cbCol.Text) && !string.IsNullOrWhiteSpace(cbAgg.Text)) {
                                string filterVal = (cbFilter.Text == "非空值 (有輸入即算)") ? "" : cbFilter.Text;
                                newCfg.Sources.Add(new DataSourceDef { DbName = cbDb.Text, TableName = cbTb.Text, ColName = cbCol.Text, FilterValue = filterVal, AggType = cbAgg.Text });
                            }
                        }
                    }

                    if (newCfg.Sources.Count == 0) { MessageBox.Show("請至少設定一組完整的資料來源！"); return; }

                    int existingIdx = _configs.FindIndex(c => c.DisplayName == newCfg.DisplayName);
                    if (existingIdx >= 0) _configs[existingIdx] = newCfg;
                    else _configs.Add(newCfg);

                    SaveSettings();
                    refreshList();
                    MessageBox.Show("儲存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                f.ShowDialog();
                _ = LoadDashboardDataAsync(); // 關閉視窗後自動重算更新畫面
            }
        }

        // =========================================================
        // 🟢 PDF 導出 (加入多區塊選擇與防呆修復)
        // =========================================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 450, Height = 450, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Top, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                
                clb.Items.Add("區間統計總計 (四大區塊)", true); 
                foreach (var kvp in _monthlyPanels) {
                    clb.Items.Add($"近三年逐月：{kvp.Key}", true);
                }
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    for (int i = 0; i < clb.Items.Count; i++) {
                        if (clb.GetItemChecked(i)) {
                            if (i == 0) selectedPanels.Add(_pnlTopBox);
                            else {
                                string key = clb.Items[i].ToString().Replace("近三年逐月：", "");
                                if (_monthlyPanels.ContainsKey(key)) selectedPanels.Add(_monthlyPanels[key]);
                            }
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

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "工安指標綜合統計表_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

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
                        pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);

                        int currentBmpIndex = 0;

                        pd.PrintPage += (s, ev) => 
                        {
                            Graphics g = ev.Graphics;
                            string headerText = $"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   查詢區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text}~{_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
                            g.DrawString(headerText, new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Brushes.Black, ev.MarginBounds.Left, ev.MarginBounds.Top - 15);

                            int currentY = ev.MarginBounds.Top + 15;
                            int bottomLimit = ev.MarginBounds.Bottom;

                            while (currentBmpIndex < bitmaps.Count) 
                            {
                                Bitmap bmp = bitmaps[currentBmpIndex];
                                float scale = (float)ev.MarginBounds.Width / bmp.Width;
                                int scaledHeight = (int)(bmp.Height * scale);

                                if (currentY + scaledHeight > bottomLimit) 
                                {
                                    if (currentY == ev.MarginBounds.Top + 15) 
                                    {
                                        scale = Math.Min(scale, (float)(bottomLimit - currentY) / bmp.Height);
                                        scaledHeight = (int)(bmp.Height * scale);
                                        g.DrawImage(bmp, ev.MarginBounds.Left, currentY, (int)(bmp.Width * scale), scaledHeight);
                                        currentY += scaledHeight + 20;
                                        currentBmpIndex++;
                                    } 
                                    else 
                                    {
                                        ev.HasMorePages = true;
                                        return;
                                    }
                                } 
                                else 
                                {
                                    g.DrawImage(bmp, ev.MarginBounds.Left, currentY, ev.MarginBounds.Width, scaledHeight);
                                    currentY += scaledHeight + 20; 
                                    currentBmpIndex++;
                                }
                            }
                            ev.HasMorePages = false;
                        };

                        pd.Print();
                        
                        // 釋放記憶體
                        foreach (var bmp in bitmaps) bmp.Dispose();

                        MessageBox.Show("PDF 匯出成功！已根據項目自動分頁並全版面 A4 對齊。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } 
                    finally 
                    { 
                        // 🟢 徹底確保滑鼠游標一定會恢復原狀
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
                    }
                }
            }
        }
    }
}
