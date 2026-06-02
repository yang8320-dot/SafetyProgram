/// FILE: Safety_System/Dashboard/App_SafetyDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
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

        // 月度/年度統計 Grid
        private DataGridView _dgvMonthly;

        // 截圖與匯出用的外框
        private Panel _pnlTopBox;
        private Panel _pnlBottomBox;

        // 設定檔路徑與快取
        private readonly string SettingsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetyDashboardSettings.txt");
        private List<SafetyConfigItem> _configs = new List<SafetyConfigItem>();

        // 定義自訂統計項目的資料結構 (支援最多3個來源)
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
            public string AggType { get; set; } // "COUNT" 或 "SUM"
        }

        public Control GetView()
        {
            LoadSettings();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            // ==========================================
            // 大框 1：區間統計與四大區塊
            // ==========================================
            _pnlTopBox = new Panel { Dock = DockStyle.Fill, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            _pnlTopBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlTopBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            // 1-1. 標題與操作列
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, Padding = new Padding(15) };
            Label lblTitle = new Label { Text = "🛡️ 工安事件與指標綜合數據看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.SteelBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes();

            Button btnSearch = new Button { Text = "🔍 查詢統計", Size = new Size(130, 32), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSearch.Click += async (s, e) => await LoadDashboardDataAsync();

            Button btnPdf = new Button { Text = "📄 導出看板 PDF", Size = new Size(180, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            Button btnSetting = new Button { Text = "⚙️ 統計項目設定", Size = new Size(180, 32), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
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

            flpTop.Controls.Add(lblTitle);
            flpTop.Controls.Add(flpControls);

            // 1-2. 四個子區塊
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
            _pnlTopBox.Controls.Add(flpTop);
            mainLayout.Controls.Add(_pnlTopBox, 0, 0);

            // ==========================================
            // 大框 2：年度逐月統計 (近三年)
            // ==========================================
            _pnlBottomBox = new Panel { Dock = DockStyle.Fill, AutoSize = true, BackColor = Color.White };
            _pnlBottomBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlBottomBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Label lblBottomTitle = new Label { Text = "📊 近三年工安指標逐月矩陣統計表", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Padding = new Padding(15, 15, 0, 15), Dock = DockStyle.Top };
            
            _dgvMonthly = new DataGridView { 
                Dock = DockStyle.Top, Height = 400, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                BorderStyle = BorderStyle.None, Margin = new Padding(15)
            };
            _dgvMonthly.EnableHeadersVisualStyles = false;
            _dgvMonthly.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkSlateBlue;
            _dgvMonthly.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvMonthly.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvMonthly.ColumnHeadersHeight = 40;
            _dgvMonthly.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvMonthly.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;

            _pnlBottomBox.Controls.Add(_dgvMonthly);
            _pnlBottomBox.Controls.Add(lblBottomTitle);
            
            mainLayout.Controls.Add(_pnlBottomBox, 0, 1);
            mainScrollPanel.Controls.Add(mainLayout);

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
                _dgvMonthly.DataSource = null;
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

            DataTable dtMonthly = new DataTable();
            dtMonthly.Columns.Add("年度", typeof(string));
            dtMonthly.Columns.Add("統計項目", typeof(string));
            for (int i = 1; i <= 12; i++) dtMonthly.Columns.Add($"{i}月", typeof(double));
            dtMonthly.Columns.Add("年度總計", typeof(double));

            await Task.Run(() =>
            {
                // 1. 計算四大區塊 (區間差異)
                dictCurr = CalculatePeriodStats(dtS, dtE);
                dictLy = CalculatePeriodStats(dtS.AddYears(-1), dtE.AddYears(-1));
                dictL2y = CalculatePeriodStats(dtS.AddYears(-2), dtE.AddYears(-2));

                // 2. 計算底部矩陣表 (近三年逐月)
                int baseYear = dtE.Year;
                int[] years = { baseYear, baseYear - 1, baseYear - 2 };

                foreach (int y in years)
                {
                    foreach (var cfg in _configs)
                    {
                        DataRow row = dtMonthly.NewRow();
                        row["年度"] = y.ToString() + "年";
                        row["統計項目"] = cfg.DisplayName;
                        double yearlyTotal = 0;

                        for (int m = 1; m <= 12; m++)
                        {
                            DateTime mStart = new DateTime(y, m, 1);
                            DateTime mEnd = new DateTime(y, m, DateTime.DaysInMonth(y, m));
                            
                            // 利用現有引擎計算單月該項目的值
                            var mResult = CalculatePeriodStats(mStart, mEnd, new List<SafetyConfigItem> { cfg });
                            double mVal = mResult.ContainsKey(cfg.DisplayName) ? mResult[cfg.DisplayName] : 0;
                            
                            row[$"{m}月"] = mVal;
                            yearlyTotal += mVal;
                        }
                        row["年度總計"] = yearlyTotal;
                        dtMonthly.Rows.Add(row);
                    }
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

                // 工安事件通常「越少越好」，所以增加為紅，減少為綠
                double diff = vCurr - vLy;
                string diffText = (diff > 0 ? "+" : "") + diff.ToString("N0") + " 件";
                Color diffColor = diff > 0 ? Color.IndianRed : (diff < 0 ? Color.ForestGreen : Color.DimGray);

                _pnlData4.Controls.Add(new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) });
            }

            // 更新 UI - 底部 Grid
            _dgvMonthly.DataSource = dtMonthly;
            _dgvMonthly.Columns["年度"].Width = 80;
            _dgvMonthly.Columns["統計項目"].Width = 180;
            _dgvMonthly.Columns["年度總計"].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            _dgvMonthly.Columns["年度總計"].DefaultCellStyle.BackColor = Color.LightYellow;

            // 將數值格式化為 N0 (整數)
            for (int i = 1; i <= 12; i++) _dgvMonthly.Columns[$"{i}月"].DefaultCellStyle.Format = "N0";
            _dgvMonthly.Columns["年度總計"].DefaultCellStyle.Format = "N0";
            
            // 處理相同年度的合併視覺效果 (透過背景顏色區隔)
            int colorToggle = 0;
            string lastYear = "";
            foreach (DataGridViewRow r in _dgvMonthly.Rows)
            {
                string curY = r.Cells["年度"].Value.ToString();
                if (curY != lastYear) { colorToggle++; lastYear = curY; }
                if (colorToggle % 2 == 0) r.DefaultCellStyle.BackColor = Color.White;
                else r.DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 250);
            }

            _dgvMonthly.ClearSelection();

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private Label CreateStatLabel(string title, double value)
        {
            return new Label { Text = $"{title}: {value:N0} 件", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        // 核心運算引擎：支援多來源聚合
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
                        // 智慧判斷日期欄位名稱
                        var cols = DataManager.GetColumnNames(src.DbName, src.TableName);
                        string dateCol = cols.Contains("日期") ? "日期" : (cols.Contains("年月") ? "年月" : (cols.Contains("年度") ? "年度" : ""));
                        if (string.IsNullOrEmpty(dateCol)) continue;

                        DataTable dt = DataManager.GetTableData(src.DbName, src.TableName, dateCol, sStr, eStr);
                        if (dt == null) continue;

                        if (src.AggType == "COUNT")
                        {
                            // 單純計算筆數 (排除刪除狀態)
                            totalVal += dt.Rows.Cast<DataRow>().Count(r => r.RowState != DataRowState.Deleted);
                        }
                        else if (src.AggType == "SUM" && !string.IsNullOrEmpty(src.ColName) && dt.Columns.Contains(src.ColName))
                        {
                            foreach (DataRow r in dt.Rows)
                            {
                                if (r.RowState == DataRowState.Deleted) continue;
                                if (double.TryParse(r[src.ColName]?.ToString().Replace(",", ""), out double v))
                                {
                                    totalVal += v;
                                }
                            }
                        }
                    }
                    catch { } // 若該資料庫/表尚未建立則安全跳過
                }
                result[cfg.DisplayName] = totalVal;
            }
            return result;
        }

        // =========================================================
        // 設定檔管理與設定視窗
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
                                if (srcParts.Length == 4)
                                {
                                    cfg.Sources.Add(new DataSourceDef { DbName = srcParts[0], TableName = srcParts[1], ColName = srcParts[2], AggType = srcParts[3] });
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
                        line += $"|{src.DbName};{src.TableName};{src.ColName};{src.AggType}";
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

            using (Form f = new Form { Text = "⚙️ 看板自訂統計項目設定", Size = new Size(1100, 700), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
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
                Label l2 = new Label { Text = "編輯 / 新增項目 (每個項目最多可聚合 3 個資料表)", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                Panel pName = new Panel { Width = 700, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 300, Location = new Point(100, 7), Font = new Font("Microsoft JhengHei UI", 12F) };
                pName.Controls.Add(txtName);
                flpEditor.Controls.Add(pName);

                // 動態生成 3 組來源選單
                var dbMap = App_DbConfig.GetDbMapCache();
                var cboDbs = new ComboBox[3];
                var cboTbs = new ComboBox[3];
                var cboCols = new ComboBox[3];
                var cboAggs = new ComboBox[3];

                for (int i = 0; i < 3; i++)
                {
                    GroupBox gb = new GroupBox { Text = $"資料來源 {i + 1}", Width = 700, Height = 100, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                    
                    cboDbs[i] = new ComboBox { Width = 150, Location = new Point(15, 45), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    cboTbs[i] = new ComboBox { Width = 200, Location = new Point(175, 45), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    cboCols[i] = new ComboBox { Width = 160, Location = new Point(385, 45), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                    cboAggs[i] = new ComboBox { Width = 100, Location = new Point(555, 45), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };

                    gb.Controls.AddRange(new Control[] { 
                        new Label { Text = "資料庫", Location = new Point(15, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cboDbs[i],
                        new Label { Text = "資料表", Location = new Point(175, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cboTbs[i],
                        new Label { Text = "計算欄位", Location = new Point(385, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cboCols[i],
                        new Label { Text = "運算方式", Location = new Point(555, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9F) }, cboAggs[i]
                    });

                    cboDbs[i].Items.Add("");
                    foreach (var kvp in dbMap) cboDbs[i].Items.Add(kvp.Key);

                    int idx = i;
                    cboDbs[idx].SelectedIndexChanged += (ss, ee) => {
                        cboTbs[idx].Items.Clear(); cboTbs[idx].Items.Add("");
                        string selDb = cboDbs[idx].Text;
                        if (!string.IsNullOrEmpty(selDb) && dbMap.ContainsKey(selDb)) {
                            foreach (var tb in dbMap[selDb].Tables.Keys) cboTbs[idx].Items.Add(tb);
                        }
                    };
                    cboTbs[idx].SelectedIndexChanged += (ss, ee) => {
                        cboCols[idx].Items.Clear(); cboCols[idx].Items.Add("Id (計數專用)");
                        string selDb = cboDbs[idx].Text; string selTb = cboTbs[idx].Text;
                        if (!string.IsNullOrEmpty(selDb) && !string.IsNullOrEmpty(selTb)) {
                            var cols = DataManager.GetColumnNames(selDb, selTb).Where(c => c != "Id");
                            foreach (var c in cols) cboCols[idx].Items.Add(c);
                        }
                    };

                    cboAggs[i].Items.AddRange(new string[] { "COUNT", "SUM" });
                    
                    flpEditor.Controls.Add(gb);
                }

                Button btnSaveRow = new Button { Text = "💾 儲存並加入清單", Width = 700, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 0), Cursor = Cursors.Hand };

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
                    var cfg = _configs[lbItems.SelectedIndex];
                    txtName.Text = cfg.DisplayName;
                    for (int i = 0; i < 3; i++) {
                        if (i < cfg.Sources.Count) {
                            cboDbs[i].Text = cfg.Sources[i].DbName;
                            cboTbs[i].Text = cfg.Sources[i].TableName;
                            cboCols[i].Text = cfg.Sources[i].ColName;
                            cboAggs[i].Text = cfg.Sources[i].AggType;
                        } else {
                            cboDbs[i].Text = ""; cboTbs[i].Text = ""; cboCols[i].Text = ""; cboAggs[i].Text = "";
                        }
                    }
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        _configs.RemoveAt(lbItems.SelectedIndex);
                        SaveSettings();
                        refreshList();
                        txtName.Clear();
                        for (int i = 0; i < 3; i++) { cboDbs[i].Text = ""; cboTbs[i].Text = ""; cboCols[i].Text = ""; cboAggs[i].Text = ""; }
                    }
                };

                btnSaveRow.Click += (ss, ee) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入顯示名稱！"); return; }
                    
                    var newCfg = new SafetyConfigItem { DisplayName = txtName.Text.Trim() };
                    for (int i = 0; i < 3; i++) {
                        if (!string.IsNullOrWhiteSpace(cboDbs[i].Text) && !string.IsNullOrWhiteSpace(cboTbs[i].Text) && !string.IsNullOrWhiteSpace(cboCols[i].Text) && !string.IsNullOrWhiteSpace(cboAggs[i].Text)) {
                            newCfg.Sources.Add(new DataSourceDef { DbName = cboDbs[i].Text, TableName = cboTbs[i].Text, ColName = cboCols[i].Text, AggType = cboAggs[i].Text });
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
        // PDF 導出
        // =========================================================
        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (_configs.Count == 0) { MessageBox.Show("無資料可導出。"); return; }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "工安指標綜合統計表_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        List<Bitmap> bitmaps = new List<Bitmap>();
                        Bitmap bmp1 = new Bitmap(_pnlTopBox.Width, _pnlTopBox.Height);
                        _pnlTopBox.DrawToBitmap(bmp1, new Rectangle(0, 0, bmp1.Width, bmp1.Height));
                        bitmaps.Add(bmp1);

                        Bitmap bmp2 = new Bitmap(_pnlBottomBox.Width, _pnlBottomBox.Height);
                        _pnlBottomBox.DrawToBitmap(bmp2, new Rectangle(0, 0, bmp2.Width, bmp2.Height));
                        bitmaps.Add(bmp2);

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
                        foreach (var bmp in bitmaps) bmp.Dispose();

                        MessageBox.Show("PDF 匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
