/// FILE: Safety_System/App_WaterDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        private Label _lblBox2Sub1, _lblBox2Sub2, _lblBox2Sub3, _lblBox2Sub4;
        private Label _lblBox3Sub1, _lblBox3Sub2, _lblBox3Sub3, _lblBox3Sub4;
        private Label _lblBox4Sub1, _lblBox4Sub2, _lblBox4Sub3, _lblBox4Sub4;

        private Panel _pnlBox2Data1, _pnlBox2Data2, _pnlBox2Data3, _pnlBox2Data4;
        private Panel _pnlBox3Data1, _pnlBox3Data2, _pnlBox3Data3, _pnlBox3Data4;
        private Panel _pnlBox4Data1, _pnlBox4Data2, _pnlBox4Data3, _pnlBox4Data4;
        private Panel _mainScrollPanel;

        private const string DbName = "Water";

        public Control GetView()
        {
            _mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };

            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1, 
                RowCount = 4 
            };

            // ==========================================
            // 大框 1：功能選單與日期查詢
            // ==========================================
            Panel box1 = new Panel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            box1.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, box1.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, Padding = new Padding(15) };
            
            Label lblTitle = new Label { Text = "💧 水資源綜合數據看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes(); 

            Button btnSearch = new Button { Text = "🔍 查詢統計", Size = new Size(130, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSearch.Click += (s, e) => LoadAllData();
            
            Button btnPdf = new Button { Text = "📄 轉存 PDF", Size = new Size(130, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearch, btnPdf
            });

            flpTop.Controls.Add(lblTitle);
            flpTop.Controls.Add(flpControls);
            box1.Controls.Add(flpTop);
            mainLayout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 建立 3 個主要數據區塊
            // ==========================================
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 水資源數據統計", Color.Teal, out _lblBox2Sub1, out _lblBox2Sub2, out _lblBox2Sub3, out _lblBox2Sub4, out _pnlBox2Data1, out _pnlBox2Data2, out _pnlBox2Data3, out _pnlBox2Data4), 0, 1);
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 回收水統計", Color.ForestGreen, out _lblBox3Sub1, out _lblBox3Sub2, out _lblBox3Sub3, out _lblBox3Sub4, out _pnlBox3Data1, out _pnlBox3Data2, out _pnlBox3Data3, out _pnlBox3Data4), 0, 2);
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 藥劑數據統計", Color.Sienna, out _lblBox4Sub1, out _lblBox4Sub2, out _lblBox4Sub3, out _lblBox4Sub4, out _pnlBox4Data1, out _pnlBox4Data2, out _pnlBox4Data3, out _pnlBox4Data4), 0, 3);

            _mainScrollPanel.Controls.Add(mainLayout);
            LoadAllData(); 
            return _mainScrollPanel;
        }

        private void InitDateComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) {
                _cboStartYear.Items.Add(i);
                _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2"));
                _cboEndMonth.Items.Add(i.ToString("D2"));
            }

            _cboStartYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboStartMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboEndYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            _cboEndMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            DateTime start = DateTime.Today.AddMonths(-1);
            DateTime end = DateTime.Today;

            SetComboValue(_cboStartYear, _cboStartMonth, _cboStartDay, start);
            SetComboValue(_cboEndYear, _cboEndMonth, _cboEndDay, end);
        }

        private void UpdateDaysCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null) return;
            int year = (int)y.SelectedItem;
            int month = int.Parse(m.SelectedItem.ToString());
            int days = DateTime.DaysInMonth(year, month);

            string currentDay = d.SelectedItem?.ToString();
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));

            if (currentDay != null && d.Items.Contains(currentDay)) {
                d.SelectedItem = currentDay;
            } else {
                d.SelectedIndex = d.Items.Count - 1; 
            }
        }

        private void SetComboValue(ComboBox y, ComboBox m, ComboBox d, DateTime date)
        {
            y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            UpdateDaysCombo(y, m, d);
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetDateFromCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            int year = (int)y.SelectedItem;
            int month = int.Parse(m.SelectedItem.ToString());
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth(year, month);
            if (day > maxDay) day = maxDay; 
            return new DateTime(year, month, day);
        }

        private Panel BuildNineGridBox(string mainTitle, Color headerColor, out Label l1, out Label l2, out Label l3, out Label l4, out Panel d1, out Panel d2, out Panel d3, out Panel d4)
        {
            Panel outer = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            outer.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, outer.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel grid = new TableLayoutPanel { 
                Dock = DockStyle.Top, AutoSize = true, 
                ColumnCount = 4, RowCount = 3, 
                Padding = new Padding(10) 
            };
            
            for (int i = 0; i < 4; i++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F)); 
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));      

            Label lblMainTitle = new Label { Text = mainTitle, Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };
            grid.Controls.Add(lblMainTitle, 0, 0);
            grid.SetColumnSpan(lblMainTitle, 4);

            l1 = CreateSubTitleLabel(headerColor);
            l2 = CreateSubTitleLabel(headerColor);
            l3 = CreateSubTitleLabel(headerColor);
            l4 = CreateSubTitleLabel(headerColor);

            grid.Controls.Add(l1, 0, 1);
            grid.Controls.Add(l2, 1, 1);
            grid.Controls.Add(l3, 2, 1);
            grid.Controls.Add(l4, 3, 1);

            d1 = CreateDataPanel(); d2 = CreateDataPanel(); d3 = CreateDataPanel(); d4 = CreateDataPanel();
            grid.Controls.Add(d1, 0, 2); grid.Controls.Add(d2, 1, 2); grid.Controls.Add(d3, 2, 2); grid.Controls.Add(d4, 3, 2);

            outer.Controls.Add(grid);
            return outer;
        }

        private Label CreateSubTitleLabel(Color bgColor)
        {
            return new Label { 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                ForeColor = Color.White, 
                BackColor = bgColor, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Dock = DockStyle.Fill, 
                Margin = new Padding(2) 
            };
        }

        private Panel CreateDataPanel()
        {
            return new FlowLayoutPanel { 
                Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 100), 
                FlowDirection = FlowDirection.TopDown, WrapContents = false, 
                BackColor = Color.FromArgb(248, 249, 250), Margin = new Padding(2), Padding = new Padding(10) 
            };
        }

        // ==========================================
        // 核心邏輯：資料運算與填寫
        // ==========================================
        private void LoadAllData()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            string dS = dtS.ToString("yyyy/MM/dd");
            string dE = dtE.ToString("yyyy/MM/dd");
            string dS_LY = dtS.AddYears(-1).ToString("yyyy/MM/dd");
            string dE_LY = dtE.AddYears(-1).ToString("yyyy/MM/dd");
            string dS_L2Y = dtS.AddYears(-2).ToString("yyyy/MM/dd");
            string dE_L2Y = dtE.AddYears(-2).ToString("yyyy/MM/dd");

            UpdateSubtitles(_lblBox2Sub1, _lblBox2Sub2, _lblBox2Sub3, _lblBox2Sub4, dS, dE, dS_LY, dE_LY, dS_L2Y, dE_L2Y, false);
            UpdateSubtitles(_lblBox3Sub1, _lblBox3Sub2, _lblBox3Sub3, _lblBox3Sub4, dS, dE, dS_LY, dE_LY, dS_L2Y, dE_L2Y, true);
            UpdateSubtitles(_lblBox4Sub1, _lblBox4Sub2, _lblBox4Sub3, _lblBox4Sub4, dS, dE, dS_LY, dE_LY, dS_L2Y, dE_L2Y, false);

            string searchS = dtS.ToString("yyyy-MM-dd");
            string searchE = dtE.ToString("yyyy-MM-dd");

            // --- 大框 2：水資源數據統計 ---
            var rawBox2_Curr = GetSumsEndingWith(searchS, searchE, "WaterMeterReadings", "WaterUsageDaily");
            var rawBox2_LY = GetSumsEndingWith(dtS.AddYears(-1).ToString("yyyy-MM-dd"), dtE.AddYears(-1).ToString("yyyy-MM-dd"), "WaterMeterReadings", "WaterUsageDaily");
            var rawBox2_L2Y = GetSumsEndingWith(dtS.AddYears(-2).ToString("yyyy-MM-dd"), dtE.AddYears(-2).ToString("yyyy-MM-dd"), "WaterMeterReadings", "WaterUsageDaily");
            
            FillDataPanels(_pnlBox2Data1, _pnlBox2Data2, _pnlBox2Data3, _pnlBox2Data4, 
                ProcessBox2Data(rawBox2_Curr), ProcessBox2Data(rawBox2_LY), ProcessBox2Data(rawBox2_L2Y));

            // --- 大框 3：回收水統計 ---
            FillDataPanels(_pnlBox3Data1, _pnlBox3Data2, _pnlBox3Data3, _pnlBox3Data4, 
                CalculateRecycleStats(searchS, searchE), CalculateRecycleStats(dtS.AddYears(-1).ToString("yyyy-MM-dd"), dtE.AddYears(-1).ToString("yyyy-MM-dd")), CalculateRecycleStats(dtS.AddYears(-2).ToString("yyyy-MM-dd"), dtE.AddYears(-2).ToString("yyyy-MM-dd")), true);

            // --- 大框 4：藥劑數據統計 ---
            var sumBox4_Curr = GetSumsEndingWith(searchS, searchE, "WaterChemicals");
            var sumBox4_LY = GetSumsEndingWith(dtS.AddYears(-1).ToString("yyyy-MM-dd"), dtE.AddYears(-1).ToString("yyyy-MM-dd"), "WaterChemicals");
            var sumBox4_L2Y = GetSumsEndingWith(dtS.AddYears(-2).ToString("yyyy-MM-dd"), dtE.AddYears(-2).ToString("yyyy-MM-dd"), "WaterChemicals");

            FillDataPanels(_pnlBox4Data1, _pnlBox4Data2, _pnlBox4Data3, _pnlBox4Data4, sumBox4_Curr, sumBox4_LY, sumBox4_L2Y);

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private void UpdateSubtitles(Label l1, Label l2, Label l3, Label l4, string s1, string e1, string s2, string e2, string s3, string e3, bool isRecycle)
        {
            string suffix = isRecycle ? "回收水量統計" : "數據統計";
            l1.Text = $"【{s1} ~ {e1}】\n區間{suffix}";
            l2.Text = $"【{s2} ~ {e2}】\n去年同期區間{suffix}";
            l3.Text = $"【{s3} ~ {e3}】\n前年同期區間{suffix}";
            l4.Text = $"【{s1} ~ {e1}】\n與去年同期差異分析"; 
        }

        // ==========================================
        // 資料客製化過濾與運算邏輯
        // ==========================================
        private Dictionary<string, double> GetSumsEndingWith(string start, string end, params string[] tableNames)
        {
            var results = new Dictionary<string, double>();
            foreach (string tbl in tableNames) {
                DataTable dt = null;
                try { dt = DataManager.GetTableData(DbName, tbl, "日期", start, end); } catch { continue; }
                if (dt == null) continue;

                // 🟢 這裡增加了 "污泥產出KG" 的特別捕捉邏輯
                var targetCols = dt.Columns.Cast<DataColumn>()
                                 .Where(c => c.ColumnName.EndsWith("日統計") || c.ColumnName == "污泥產出KG")
                                 .Select(c => c.ColumnName).ToList();

                foreach (DataRow r in dt.Rows) {
                    foreach (string col in targetCols) {
                        // 如果是污泥產出KG，名字會直接叫做 污泥量
                        string cleanName = col == "污泥產出KG" ? "污泥量" : col.Replace("日統計", ""); 
                        if (!results.ContainsKey(cleanName)) results[cleanName] = 0;
                        if (double.TryParse(r[col]?.ToString().Replace(",", ""), out double v)) {
                            results[cleanName] += v;
                        }
                    }
                }
            }
            return results;
        }

        private Dictionary<string, double> ProcessBox2Data(Dictionary<string, double> raw)
        {
            var res = new Dictionary<string, double>();
            foreach (var kvp in raw)
            {
                // 需求 2.1, 2.3 & 2: 隱藏回收水6吋、軟水A/B/C、回收水雙介質A/B
                if (kvp.Key.Contains("回收水6吋") || kvp.Key.Contains("軟水") || kvp.Key.Contains("回收水雙介質")) continue;

                // 需求 2.5: 用電量 * 100
                double val = kvp.Key == "用電量" ? kvp.Value * 100 : kvp.Value;
                res[kvp.Key] = val;

                // 需求 2.4: 在濃縮水至逆洗池後面加入濃縮水合計
                if (kvp.Key == "濃縮水至逆洗池" || (kvp.Key == "濃縮水至冷卻水池" && !raw.ContainsKey("濃縮水至逆洗池"))) {
                    res["濃縮水合計"] = (raw.ContainsKey("濃縮水至冷卻水池") ? raw["濃縮水至冷卻水池"] : 0) + 
                                        (raw.ContainsKey("濃縮水至逆洗池") ? raw["濃縮水至逆洗池"] : 0);
                }
            }

            // 需求 3 & 2: 計算 總回收水量
            if (raw.ContainsKey("回收水雙介質A") || raw.ContainsKey("回收水雙介質B")) {
                res["總回收水量"] = (raw.ContainsKey("回收水雙介質A") ? raw["回收水雙介質A"] : 0) + 
                                    (raw.ContainsKey("回收水雙介質B") ? raw["回收水雙介質B"] : 0);
            }
            return res;
        }

        private Dictionary<string, double> CalculateRecycleStats(string start, string end)
        {
            // 需求 2: 隱藏雙介質 A 與 B，僅顯示廢水處理量、總回收水量、回收率
            var dict = new Dictionary<string, double> {
                { "廢水處理量", 0 }, { "總回收水量", 0 }, { "回收率(%)", 0 }
            };

            double sumA = 0;
            double sumB = 0;

            DataTable dt = null;
            try { dt = DataManager.GetTableData(DbName, "WaterMeterReadings", "日期", start, end); } catch { return dict; }
            if (dt == null) return dict;

            foreach (DataRow r in dt.Rows) {
                if (dt.Columns.Contains("廢水處理量日統計") && double.TryParse(r["廢水處理量日統計"]?.ToString().Replace(",", ""), out double w)) dict["廢水處理量"] += w;
                if (dt.Columns.Contains("回收水雙介質A日統計") && double.TryParse(r["回收水雙介質A日統計"]?.ToString().Replace(",", ""), out double a)) sumA += a;
                if (dt.Columns.Contains("回收水雙介質B日統計") && double.TryParse(r["回收水雙介質B日統計"]?.ToString().Replace(",", ""), out double b)) sumB += b;
            }

            dict["總回收水量"] = sumA + sumB;
            if (dict["廢水處理量"] > 0) {
                dict["回收率(%)"] = (dict["總回收水量"] / dict["廢水處理量"]) * 100;
            }

            return dict;
        }

        // ==========================================
        // 渲染 Label 數據與差異顏色處理
        // ==========================================
        private void FillDataPanels(Panel p1, Panel p2, Panel p3, Panel p4, Dictionary<string, double> curr, Dictionary<string, double> ly, Dictionary<string, double> l2y, bool isRecycleRate = false)
        {
            p1.Controls.Clear(); p2.Controls.Clear(); p3.Controls.Clear(); p4.Controls.Clear();

            foreach (var kvp in curr)
            {
                string key = kvp.Key;
                double vCurr = kvp.Value;
                double vLy = ly.ContainsKey(key) ? ly[key] : 0;
                double vL2y = l2y.ContainsKey(key) ? l2y[key] : 0;

                p1.Controls.Add(CreateStatLabel(key, vCurr));
                p2.Controls.Add(CreateStatLabel(key, vLy));
                p3.Controls.Add(CreateStatLabel(key, vL2y));

                string diffText = "無基期";
                Color diffColor = Color.DimGray;

                if (vLy > 0) {
                    double yoy = ((vCurr - vLy) / vLy) * 100;
                    if (isRecycleRate && key.Contains("回收率")) yoy = vCurr - vLy; 

                    string formatStr = key.Contains("回收率") ? "N1" : "N0";
                    diffText = (yoy > 0 ? "+" : "") + yoy.ToString(formatStr) + " %";
                    
                    diffColor = yoy > 0 ? Color.IndianRed : (yoy < 0 ? Color.ForestGreen : Color.DimGray); 
                } else if (vCurr > 0) {
                    diffText = "新數據";
                    diffColor = Color.SteelBlue;
                }

                Label lblDiff = new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
                p4.Controls.Add(lblDiff);
            }
        }

        private Label CreateStatLabel(string title, double value)
        {
            string unit = " M3";
            if (title.Contains("用電")) unit = " KWH";
            else if (title.Contains("%") || title.Contains("率")) unit = " %";
            else if (title.Contains("包")) unit = " 包";
            else if (title.Contains("污泥量")) unit = " KG"; // 🟢 污泥量專屬單位

            string format = title.Contains("回收率") ? "N1" : "N0";

            return new Label { 
                Text = $"{title}: {value.ToString(format)}{unit}", 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                ForeColor = Color.FromArgb(45,45,45), 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 8) 
            };
        }

        // ==========================================
        // 匯出 PDF 功能：A4 橫向、自動縮放、加上導出時間
        // ==========================================
        private void BtnPdf_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "水資源綜合統計表_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        int originalHeight = _mainScrollPanel.Height;
                        _mainScrollPanel.Height = _mainScrollPanel.DisplayRectangle.Height; 
                        
                        Bitmap bmp = new Bitmap(_mainScrollPanel.Width, _mainScrollPanel.Height);
                        _mainScrollPanel.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                        
                        _mainScrollPanel.Height = originalHeight; 

                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        
                        pd.DefaultPageSettings.Landscape = true; 
                        pd.DefaultPageSettings.Margins = new Margins(20, 20, 30, 30);

                        int currentY = 0;

                        pd.PrintPage += (s, ev) => {
                            Graphics g = ev.Graphics;
                            
                            string headerText = $"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   查詢區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text} ~ {_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
                            Font fontHeader = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                            g.DrawString(headerText, fontHeader, Brushes.Black, ev.MarginBounds.Left, ev.MarginBounds.Top - 20);

                            int printTop = ev.MarginBounds.Top + 15;
                            int printHeight = ev.MarginBounds.Height - 15;

                            float scale = (float)ev.MarginBounds.Width / bmp.Width;
                            int sourceHeightFit = (int)(printHeight / scale);

                            Rectangle destRect = new Rectangle(ev.MarginBounds.Left, printTop, ev.MarginBounds.Width, printHeight);
                            Rectangle srcRect = new Rectangle(0, currentY, bmp.Width, sourceHeightFit);

                            if (currentY + sourceHeightFit > bmp.Height) {
                                srcRect.Height = bmp.Height - currentY;
                                destRect.Height = (int)(srcRect.Height * scale);
                            }

                            g.DrawImage(bmp, destRect, srcRect, GraphicsUnit.Pixel);
                            currentY += sourceHeightFit;

                            ev.HasMorePages = currentY < bmp.Height;
                        };

                        pd.Print();
                        bmp.Dispose();

                        MessageBox.Show("PDF 匯出成功！已縮放至全版面 A4。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }
    }
}
