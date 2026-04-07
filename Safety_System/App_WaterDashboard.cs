using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        private ComboBox _cboYear, _cboMonth, _cboMetric;
        private Label _lblKpi1, _lblKpi2, _lblKpi3;
        private Chart _chartDaily, _chartMonthly, _chartYearly;
        private const string DbName = "Water";

        public Control GetView()
        {
            Panel mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true };

            // 1. 頂部控制列 (篩選條件)
            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = Color.White };
            Label lblTitle = new Label { Text = "💧 水資源管理與分析儀表板", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), AutoSize = true, Location = new Point(20, 20), ForeColor = Color.DarkSlateBlue };
            
            FlowLayoutPanel flpFilters = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Padding = new Padding(10, 25, 20, 0) };
            
            _cboYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboMonth = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboMetric = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            int currYear = DateTime.Now.Year;
            for (int i = currYear - 10; i <= currYear; i++) _cboYear.Items.Add(i);
            for (int i = 1; i <= 12; i++) _cboMonth.Items.Add(i.ToString("D2"));
            
            _cboYear.SelectedItem = currYear;
            _cboMonth.SelectedItem = DateTime.Now.Month.ToString("D2");

            _cboMetric.Items.Add("廠區自來水量 (用水)");
            _cboMetric.Items.Add("廢水處理量 (水站)");
            _cboMetric.Items.Add("PAC用藥量 (水站)");
            _cboMetric.SelectedIndex = 0;

            Button btnRefresh = new Button { Text = "📊 產生報表", Size = new Size(120, 32), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
            btnRefresh.Click += (s, e) => LoadDashboardData();

            flpFilters.Controls.AddRange(new Control[] { 
                new Label { Text = "分析標的:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _cboMetric,
                new Label { Text = "基準年月:", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _cboYear, new Label { Text = "年", AutoSize=true, Margin=new Padding(0,5,5,0) }, _cboMonth, new Label { Text = "月", AutoSize=true, Margin=new Padding(0,5,15,0) },
                btnRefresh 
            });

            pnlTop.Controls.Add(lblTitle);
            pnlTop.Controls.Add(flpFilters);
            mainPanel.Controls.Add(pnlTop);

            // 2. KPI 摘要卡片區
            FlowLayoutPanel flpKpis = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(20, 15, 20, 10), BackColor = Color.WhiteSmoke };
            _lblKpi1 = CreateKpiCard(flpKpis, "本月累計", "0", Color.Teal);
            _lblKpi2 = CreateKpiCard(flpKpis, "去年同月累計", "0", Color.DimGray);
            _lblKpi3 = CreateKpiCard(flpKpis, "YoY 成長率", "0%", Color.DarkOrange);
            mainPanel.Controls.Add(flpKpis);

            // 3. 圖表區
            TableLayoutPanel tlpCharts = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(15) };
            tlpCharts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
            tlpCharts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            tlpCharts.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tlpCharts.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

            _chartDaily = CreateChart("每日變化比較 (本月 vs 去年同月)");
            _chartMonthly = CreateChart("每月趨勢比較 (近三年)");
            _chartYearly = CreateChart("年度總量比較 (近三年)");

            tlpCharts.Controls.Add(_chartDaily, 0, 0); 
            tlpCharts.SetRowSpan(_chartDaily, 2);      
            tlpCharts.Controls.Add(_chartMonthly, 1, 0); 
            tlpCharts.Controls.Add(_chartYearly, 1, 1);  
            mainPanel.Controls.Add(tlpCharts);

            LoadDashboardData();
            return mainPanel;
        }

        private Label CreateKpiCard(FlowLayoutPanel parent, string title, string defaultVal, Color color)
        {
            Panel card = new Panel { Size = new Size(250, 70), BackColor = Color.White, Margin = new Padding(0, 0, 20, 0) };
            card.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, card.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            Label lblT = new Label { Text = title, Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.Gray, Location = new Point(15, 10), AutoSize = true };
            Label lblV = new Label { Text = defaultVal, Font = new Font("Arial", 20F, FontStyle.Bold), ForeColor = color, Location = new Point(15, 30), AutoSize = true };
            card.Controls.Add(lblT); card.Controls.Add(lblV);
            parent.Controls.Add(card);
            return lblV;
        }

        private Chart CreateChart(string title)
        {
            Chart chart = new Chart { Dock = DockStyle.Fill, BackColor = Color.White };
            ChartArea ca = new ChartArea { Name = "MainArea", BackColor = Color.WhiteSmoke };
            ca.AxisX.MajorGrid.LineColor = Color.Gainsboro; ca.AxisY.MajorGrid.LineColor = Color.Gainsboro;
            chart.ChartAreas.Add(ca);
            chart.Titles.Add(new Title(title, Docking.Top, new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Color.DarkSlateGray));
            chart.Legends.Add(new Legend { Docking = Docking.Bottom, Alignment = StringAlignment.Center });
            return chart;
        }

        private DataTable SafeGetTableData(string dbName, string tableName, string startDate, string endDate)
        {
            try { return DataManager.GetTableData(dbName, tableName, "日期", startDate, endDate); }
            catch { return new DataTable(); }
        }

        private void LoadDashboardData()
        {
            if (_cboYear.SelectedItem == null || _cboMonth.SelectedItem == null || _cboMetric.SelectedItem == null) return;
            int targetYear = (int)_cboYear.SelectedItem;
            int targetMonth = int.Parse(_cboMonth.SelectedItem.ToString());
            string metric = _cboMetric.SelectedItem.ToString();
            string tableName = ""; string valCol = "";

            if (metric.Contains("廠區自來水量")) { tableName = "WaterUsageDaily"; valCol = "廠區自來水量日統計"; }
            else if (metric.Contains("廢水處理量")) { tableName = "WaterMeterReadings"; valCol = "廢水處理量日統計"; }
            else if (metric.Contains("PAC")) { tableName = "WaterChemicals"; valCol = "PAC日統計"; }

            DataTable dtY0 = SafeGetTableData(DbName, tableName, $"{targetYear}-01-01", $"{targetYear}-12-31");
            DataTable dtY1 = SafeGetTableData(DbName, tableName, $"{targetYear - 1}-01-01", $"{targetYear - 1}-12-31");
            DataTable dtY2 = SafeGetTableData(DbName, tableName, $"{targetYear - 2}-01-01", $"{targetYear - 2}-12-31");

            DrawDailyChart(dtY0, dtY1, targetYear, targetMonth, valCol);
            DrawMonthlyChart(dtY0, dtY1, dtY2, targetYear, valCol);
            DrawYearlyChart(dtY0, dtY1, dtY2, targetYear, valCol);
            UpdateKpis(dtY0, dtY1, targetMonth, valCol);
        }

        private void UpdateKpis(DataTable dtY0, DataTable dtY1, int month, string valCol)
        {
            double sumY0 = GetSumForMonth(dtY0, month, valCol);
            double sumY1 = GetSumForMonth(dtY1, month, valCol);
            _lblKpi1.Text = sumY0.ToString("N1");
            _lblKpi2.Text = sumY1.ToString("N1");

            if (sumY1 > 0) {
                double yoy = ((sumY0 - sumY1) / sumY1) * 100;
                _lblKpi3.Text = (yoy > 0 ? "+" : "") + yoy.ToString("N1") + " %";
                _lblKpi3.ForeColor = yoy > 0 ? Color.IndianRed : Color.Green;
            } else if (sumY0 > 0) {
                _lblKpi3.Text = "新數據"; _lblKpi3.ForeColor = Color.SteelBlue;
            } else {
                _lblKpi3.Text = "無基期"; _lblKpi3.ForeColor = Color.DimGray;
            }
        }

        private void DrawDailyChart(DataTable dtY0, DataTable dtY1, int year, int month, string valCol)
        {
            _chartDaily.Series.Clear();
            int days = DateTime.DaysInMonth(year, month);
            Series s0 = new Series($"{year}年{month}月") { ChartType = SeriesChartType.Line, BorderWidth = 3, Color = Color.SteelBlue };
            Series s1 = new Series($"{year - 1}年{month}月") { ChartType = SeriesChartType.Line, BorderWidth = 2, Color = Color.Silver, BorderDashStyle = ChartDashStyle.Dash };

            for (int day = 1; day <= days; day++) {
                s0.Points.AddXY(day, GetValForDay(dtY0, month, day, valCol));
                s1.Points.AddXY(day, GetValForDay(dtY1, month, day, valCol));
            }
            _chartDaily.Series.Add(s0); _chartDaily.Series.Add(s1);
            _chartDaily.ChartAreas[0].AxisX.Interval = 2;
        }

        private void DrawMonthlyChart(DataTable dtY0, DataTable dtY1, DataTable dtY2, int year, string valCol)
        {
            _chartMonthly.Series.Clear();
            Series s0 = new Series($"{year}年") { ChartType = SeriesChartType.Column, Color = Color.SteelBlue };
            Series s1 = new Series($"{year - 1}年") { ChartType = SeriesChartType.Column, Color = Color.Teal };
            Series s2 = new Series($"{year - 2}年") { ChartType = SeriesChartType.Column, Color = Color.LightGray };

            for (int m = 1; m <= 12; m++) {
                s0.Points.AddXY($"{m}月", GetSumForMonth(dtY0, m, valCol));
                s1.Points.AddXY($"{m}月", GetSumForMonth(dtY1, m, valCol));
                s2.Points.AddXY($"{m}月", GetSumForMonth(dtY2, m, valCol));
            }
            _chartMonthly.Series.Add(s2); _chartMonthly.Series.Add(s1); _chartMonthly.Series.Add(s0);
        }

        private void DrawYearlyChart(DataTable dtY0, DataTable dtY1, DataTable dtY2, int year, string valCol)
        {
            _chartYearly.Series.Clear();
            Series s = new Series("年度總量") { ChartType = SeriesChartType.Bar, IsValueShownAsLabel = true, LabelFormat = "N1" };
            s.Points.AddXY($"{year - 2}年", GetSumForYear(dtY2, valCol));
            s.Points.AddXY($"{year - 1}年", GetSumForYear(dtY1, valCol));
            s.Points.AddXY($"{year}年", GetSumForYear(dtY0, valCol));
            _chartYearly.Series.Add(s);
        }

        // ==========================================
        // 使用傳統 foreach 迭代，避免建置時 AsEnumerable() 解析失敗
        // ==========================================
        private double GetValForDay(DataTable dt, int month, int day, string valCol)
        {
            if (dt == null || dt.Rows.Count == 0 || !dt.Columns.Contains(valCol)) return 0;
            string target = string.Format("-{0:D2}-{1:D2}", month, day);
            
            foreach (DataRow row in dt.Rows) {
                if (row["日期"] != null && row["日期"].ToString().Contains(target)) {
                    double v;
                    if (double.TryParse(row[valCol].ToString().Replace(",", "").Trim(), out v)) return v;
                }
            }
            return 0;
        }

        private double GetSumForMonth(DataTable dt, int month, string valCol)
        {
            if (dt == null || dt.Rows.Count == 0 || !dt.Columns.Contains(valCol)) return 0;
            string target = string.Format("-{0:D2}-", month);
            double sum = 0;

            foreach (DataRow row in dt.Rows) {
                if (row["日期"] != null && row["日期"].ToString().Contains(target)) {
                    double v;
                    if (double.TryParse(row[valCol].ToString().Replace(",", "").Trim(), out v)) sum += v;
                }
            }
            return sum;
        }

        private double GetSumForYear(DataTable dt, string valCol)
        {
            if (dt == null || dt.Rows.Count == 0 || !dt.Columns.Contains(valCol)) return 0;
            double sum = 0;

            foreach (DataRow row in dt.Rows) {
                double v;
                if (double.TryParse(row[valCol].ToString().Replace(",", "").Trim(), out v)) sum += v;
            }
            return sum;
        }
    }
}
