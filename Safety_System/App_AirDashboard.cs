/// FILE: Safety_System/App_AirDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Safety_System
{
    public class App_AirDashboard
    {
        // 頂部控制項 (空污費)
        private ComboBox _cboAirYear, _cboAirQuarter;
        private Label _lblAirEmissionsCurr, _lblAirEmissionsLY, _lblAirEmissionsL2Y, _lblAirEmissionsDiff;
        private Label _lblAirFeeCurr, _lblAirFeeLY, _lblAirFeeL2Y, _lblAirFeeDiff;
        private Chart _airChart;

        // 底部控制項 (原物料統計)
        private ComboBox _cboMatStartYear, _cboMatStartMonth, _cboMatEndYear, _cboMatEndMonth;
        private DataGridView _dgvMaterial;

        // 原物料設定模型
        private class MatConfig 
        {
            public string Alias { get; set; }
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public double Multiplier { get; set; }
        }
        private List<MatConfig> _matConfigs = new List<MatConfig>();
        private readonly string ConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AirMaterialSettings.txt");

        // 定義可選的資料庫清單
        private readonly Dictionary<string, string> _dbMap = new Dictionary<string, string> {
            { "Water", "水污" }, { "Air", "空污" }, { "Waste", "廢棄物及產能" }, 
            { "Chemical", "化學品" }, { "Fire", "消防" }, { "Safety", "工安" }, { "Purchase", "請購" }
        };

        public Control GetView()
        {
            LoadMaterialConfigs();

            Panel mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 4 };

            // ==========================================
            // 第一區塊：空污費查詢與分析 (Part 1 & 2)
            // ==========================================
            GroupBox boxAir = new GroupBox { Text = "☁️ 台灣玻璃彰濱廠 - 空污費申報【排放量】統計", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0, 0, 0, 20), BackColor = Color.White };
            
            FlowLayoutPanel flpAirFilter = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 0, 0, 15) };
            _cboAirYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100, Margin = new Padding(0, 4, 10, 0) };
            _cboAirQuarter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Margin = new Padding(0, 4, 10, 0) };
            
            int currYear = DateTime.Today.Year;
            for (int i = currYear - 10; i <= currYear; i++) _cboAirYear.Items.Add(i.ToString());
            _cboAirYear.SelectedItem = currYear.ToString();

            _cboAirQuarter.Items.AddRange(new string[] { "全年 (Q1~Q4)", "第一季", "第二季", "第三季", "第四季" });
            _cboAirQuarter.SelectedIndex = 0;

            Button btnSearchAir = new Button { Text = "🔍 查詢空污統計", Size = new Size(160, 35), BackColor = Color.DeepSkyBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSearchAir.FlatAppearance.BorderSize = 0;
            btnSearchAir.Click += (s, e) => LoadAirPollutionData();

            flpAirFilter.Controls.AddRange(new Control[] {
                new Label { Text = "查詢年度:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) }, _cboAirYear,
                new Label { Text = "申報季度:", AutoSize = true, Margin = new Padding(15, 8, 5, 0) }, _cboAirQuarter,
                btnSearchAir
            });

            // 數據方塊區
            TableLayoutPanel tlpAirData = new TableLayoutPanel { Dock = DockStyle.Top, Height = 100, ColumnCount = 4, RowCount = 2, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single };
            for (int i = 0; i < 4; i++) tlpAirData.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpAirData.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            tlpAirData.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            string[] airHeaders = { "當期申報數據", "去年同期數據", "前年同期數據", "與去年同期差異" };
            for (int i = 0; i < 4; i++) {
                tlpAirData.Controls.Add(new Label { Text = airHeaders[i], Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.DeepSkyBlue, ForeColor = Color.White }, i, 0);
            }

            _lblAirEmissionsCurr = CreateDataLabel(); _lblAirFeeCurr = CreateDataLabel();
            _lblAirEmissionsLY = CreateDataLabel();   _lblAirFeeLY = CreateDataLabel();
            _lblAirEmissionsL2Y = CreateDataLabel();  _lblAirFeeL2Y = CreateDataLabel();
            _lblAirEmissionsDiff = CreateDataLabel(); _lblAirFeeDiff = CreateDataLabel();

            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsCurr, _lblAirFeeCurr), 0, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsLY, _lblAirFeeLY), 1, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsL2Y, _lblAirFeeL2Y), 2, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsDiff, _lblAirFeeDiff), 3, 1);

            // 圖表區
            _airChart = new Chart { Dock = DockStyle.Top, Height = 350, Margin = new Padding(0, 15, 0, 0) };
            ChartArea ca = new ChartArea("MainArea");
            ca.AxisX.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.Title = "排放量 (KG/Ton)";
            ca.AxisY2.Title = "繳費金額 (NTD)";
            ca.AxisY2.MajorGrid.Enabled = false;
            _airChart.ChartAreas.Add(ca);
            _airChart.Legends.Add(new Legend("Legend1") { Docking = Docking.Top, Alignment = StringAlignment.Center });

            boxAir.Controls.Add(_airChart);
            boxAir.Controls.Add(tlpAirData);
            boxAir.Controls.Add(flpAirFilter);
            tlpMain.Controls.Add(boxAir, 0, 0);

            // ==========================================
            // 第二區塊：原物料使用紀錄統計表 (Part 3)
            // ==========================================
            GroupBox boxMaterial = new GroupBox { Text = "📊 台灣玻璃彰濱廠 - 原物料使用紀錄統計表", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), BackColor = Color.White };
            
            FlowLayoutPanel flpMatFilter = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 0, 0, 15), WrapContents = false };
            
            _cboMatStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80, Margin = new Padding(0, 4, 5, 0) };
            _cboMatStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = new Padding(0, 4, 10, 0) };
            _cboMatEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80, Margin = new Padding(0, 4, 5, 0) };
            _cboMatEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = new Padding(0, 4, 15, 0) };

            for (int i = currYear - 10; i <= currYear; i++) {
                _cboMatStartYear.Items.Add(i.ToString()); _cboMatEndYear.Items.Add(i.ToString());
            }
            for (int i = 1; i <= 12; i++) {
                _cboMatStartMonth.Items.Add(i.ToString("D2")); _cboMatEndMonth.Items.Add(i.ToString("D2"));
            }

            DateTime lastMonth = DateTime.Today.AddMonths(-1);
            _cboMatStartYear.SelectedItem = lastMonth.Year.ToString();
            _cboMatStartMonth.SelectedItem = lastMonth.Month.ToString("D2");
            _cboMatEndYear.SelectedItem = DateTime.Today.Year.ToString();
            _cboMatEndMonth.SelectedItem = DateTime.Today.Month.ToString("D2");

            Button btnSearchMat = new Button { Text = "🔍 查詢原物料統計", Size = new Size(170, 35), BackColor = Color.SeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSearchAir.FlatAppearance.BorderSize = 0;
            btnSearchMat.Click += async (s, e) => await LoadMaterialDataAsync();

            Button btnConfigMat = new Button { Text = "⚙️ 設定查詢欄位", Size = new Size(160, 35), BackColor = Color.DimGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnConfigMat.Click += (s, e) => {
                OpenMaterialConfigDialog();
                _ = LoadMaterialDataAsync();
            };

            flpMatFilter.Controls.AddRange(new Control[] {
                new Label { Text = "年月區間:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) }, 
                _cboMatStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _cboMatStartMonth, new Label { Text = "月 ~", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                _cboMatEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _cboMatEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 15, 0) },
                btnSearchMat, btnConfigMat
            });

            _dgvMaterial = new DataGridView { 
                Dock = DockStyle.Top, Height = 400, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 10, 0, 0)
            };
            _dgvMaterial.EnableHeadersVisualStyles = false;
            _dgvMaterial.ColumnHeadersDefaultCellStyle.BackColor = Color.SeaGreen;
            _dgvMaterial.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvMaterial.ColumnHeadersHeight = 40;
            _dgvMaterial.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 250, 245);

            boxMaterial.Controls.Add(_dgvMaterial);
            boxMaterial.Controls.Add(flpMatFilter);
            tlpMain.Controls.Add(boxMaterial, 0, 1);

            mainPanel.Controls.Add(tlpMain);

            LoadAirPollutionData();
            _ = LoadMaterialDataAsync();

            return mainPanel;
        }

        // ====================================================
        // Part 1 & 2: 空污費邏輯
        // ====================================================
        private Label CreateDataLabel() => new Label { Font = new Font("Microsoft JhengHei UI", 11F), ForeColor = Color.FromArgb(50, 50, 50), AutoSize = true, Margin = new Padding(0, 5, 0, 5) };

        private FlowLayoutPanel CreateDataCell(Label l1, Label l2)
        {
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, Padding = new Padding(10) };
            flp.Controls.Add(l1);
            flp.Controls.Add(l2);
            return flp;
        }

        private void LoadAirPollutionData()
        {
            string yearStr = _cboAirYear.SelectedItem.ToString();
            string qtrStr = _cboAirQuarter.SelectedItem.ToString();
            int year = int.Parse(yearStr);

            var currData = FetchAirData(year, qtrStr);
            var lyData = FetchAirData(year - 1, qtrStr);
            var l2yData = FetchAirData(year - 2, qtrStr);

            UpdateAirLabels(_lblAirEmissionsCurr, _lblAirFeeCurr, currData);
            UpdateAirLabels(_lblAirEmissionsLY, _lblAirFeeLY, lyData);
            UpdateAirLabels(_lblAirEmissionsL2Y, _lblAirFeeL2Y, l2yData);

            _lblAirEmissionsDiff.Text = $"排放量: {CalculateDiff(currData.Emissions, lyData.Emissions)}";
            _lblAirEmissionsDiff.ForeColor = currData.Emissions > lyData.Emissions ? Color.IndianRed : Color.ForestGreen;
            
            _lblAirFeeDiff.Text = $"繳費金額: {CalculateDiff(currData.Fee, lyData.Fee)}";
            _lblAirFeeDiff.ForeColor = currData.Fee > lyData.Fee ? Color.IndianRed : Color.ForestGreen;

            UpdateAirChart(year, currData, lyData, l2yData);
        }

        private (double Emissions, double Fee) FetchAirData(int year, string quarterMode)
        {
            double emissions = 0, fee = 0;
            try {
                DataTable dt = DataManager.GetTableData("Air", "AirPollution", "", "", "");
                if (dt != null) {
                    foreach (DataRow r in dt.Rows) {
                        string dbYear = r["年度"]?.ToString().Trim();
                        string dbQtr = r["季度"]?.ToString().Trim();

                        if (dbYear == year.ToString()) {
                            bool matchQtr = quarterMode.Contains("全年") || quarterMode == dbQtr;
                            if (matchQtr) {
                                if (double.TryParse(r["排放量"]?.ToString().Replace(",", ""), out double em)) emissions += em;
                                if (double.TryParse(r["繳費金額"]?.ToString().Replace(",", ""), out double f)) fee += f;
                            }
                        }
                    }
                }
            } catch { }
            return (emissions, fee);
        }

        private void UpdateAirLabels(Label lEmissions, Label lFee, (double Emissions, double Fee) data)
        {
            lEmissions.Text = $"排放總量: {data.Emissions:N2}";
            lFee.Text = $"繳費金額: $ {data.Fee:N0}";
        }

        private string CalculateDiff(double curr, double ly)
        {
            if (ly == 0) return "無基期";
            double diff = ((curr - ly) / ly) * 100;
            return (diff > 0 ? "+" : "") + diff.ToString("N1") + " %";
        }

        private void UpdateAirChart(int baseYear, (double E, double F) curr, (double E, double F) ly, (double E, double F) l2y)
        {
            _airChart.Series.Clear();

            Series sEmissions = new Series("排放量") { ChartType = SeriesChartType.Column, YAxisType = AxisType.Primary, Color = Color.SteelBlue, IsValueShownAsLabel = true };
            Series sFee = new Series("繳費金額") { ChartType = SeriesChartType.Line, BorderWidth = 3, MarkerStyle = MarkerStyle.Circle, MarkerSize = 8, YAxisType = AxisType.Secondary, Color = Color.IndianRed, IsValueShownAsLabel = true };

            sEmissions.Points.AddXY((baseYear - 2).ToString(), l2y.E);
            sEmissions.Points.AddXY((baseYear - 1).ToString(), ly.E);
            sEmissions.Points.AddXY(baseYear.ToString(), curr.E);

            sFee.Points.AddXY((baseYear - 2).ToString(), l2y.F);
            sFee.Points.AddXY((baseYear - 1).ToString(), ly.F);
            sFee.Points.AddXY(baseYear.ToString(), curr.F);

            _airChart.Series.Add(sEmissions);
            _airChart.Series.Add(sFee);
            _airChart.DataBind();
        }

        // ====================================================
        // Part 3: 原物料設定與運算邏輯
        // ====================================================
        private void LoadMaterialConfigs()
        {
            _matConfigs.Clear();
            if (File.Exists(ConfigFile)) {
                try {
                    foreach (var line in File.ReadAllLines(ConfigFile, Encoding.UTF8)) {
                        var p = line.Split('|');
                        if (p.Length >= 5) {
                            _matConfigs.Add(new MatConfig { Alias = p[0], DbName = p[1], TableName = p[2], ColName = p[3], Multiplier = double.Parse(p[4]) });
                        }
                    }
                } catch { }
            }
        }

        private void SaveMaterialConfigs()
        {
            try {
                var lines = _matConfigs.Select(c => $"{c.Alias}|{c.DbName}|{c.TableName}|{c.ColName}|{c.Multiplier}").ToArray();
                File.WriteAllLines(ConfigFile, lines, Encoding.UTF8);
            } catch { }
        }

        private void OpenMaterialConfigDialog()
        {
            using (Form f = new Form { Text = "⚙️ 設定原物料查詢組合 (最多5組)", Size = new Size(1100, 480), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 6, Padding = new Padding(15) };
                
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F)); // 名稱
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F)); // 庫
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F)); // 表
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F)); // 欄位
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));  // 換算

                string[] headers = { "顯示名稱 (如: 柴油(公秉))", "來源資料庫", "來源資料表", "加總欄位", "單位換算規則" };
                for (int i = 0; i < 5; i++) tlp.Controls.Add(new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) }, i, 0);

                var rowsUi = new List<(TextBox txtName, ComboBox cbDb, ComboBox cbTb, ComboBox cbCol, ComboBox cbConv)>();

                for (int i = 0; i < 5; i++)
                {
                    TextBox txtName = new TextBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbTb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbConv = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

                    foreach (var k in _dbMap.Keys) cbDb.Items.Add(k);

                    cbConv.Items.AddRange(new string[] { "無換算 (x1)", "公噸 ➔ 公斤 (x1000)", "公斤 ➔ 公噸 (x0.001)", "公升 ➔ 公秉 (x0.001)", "公秉 ➔ 公升 (x1000)" });

                    // 綁定連動事件
                    cbDb.SelectedIndexChanged += (s, e) => {
                        cbTb.Items.Clear(); cbCol.Items.Clear();
                        if (cbDb.SelectedItem != null) {
                            var tbs = GetTablesForDb(cbDb.SelectedItem.ToString());
                            cbTb.Items.AddRange(tbs.ToArray());
                        }
                    };

                    cbTb.SelectedIndexChanged += (s, e) => {
                        cbCol.Items.Clear();
                        if (cbDb.SelectedItem != null && cbTb.SelectedItem != null) {
                            var cols = DataManager.GetColumnNames(cbDb.SelectedItem.ToString(), cbTb.SelectedItem.ToString());
                            foreach(var c in cols) {
                                if (c != "Id" && c != "日期" && c != "年月" && c != "年度" && c != "備註" && c != "附件檔案") {
                                    cbCol.Items.Add(c);
                                }
                            }
                        }
                    };

                    // 填入既有設定
                    if (i < _matConfigs.Count) {
                        var conf = _matConfigs[i];
                        txtName.Text = conf.Alias;
                        if (cbDb.Items.Contains(conf.DbName)) cbDb.SelectedItem = conf.DbName;
                        if (cbTb.Items.Contains(conf.TableName)) cbTb.SelectedItem = conf.TableName;
                        if (cbCol.Items.Contains(conf.ColName)) cbCol.SelectedItem = conf.ColName;
                        
                        if (conf.Multiplier == 1000) cbConv.SelectedIndex = 1; // 假設是 x1000
                        else if (conf.Multiplier == 0.001) cbConv.SelectedIndex = 2; // 假設是 x0.001
                        else cbConv.SelectedIndex = 0;
                    } else {
                        cbConv.SelectedIndex = 0;
                    }

                    tlp.Controls.Add(txtName, 0, i + 1);
                    tlp.Controls.Add(cbDb, 1, i + 1);
                    tlp.Controls.Add(cbTb, 2, i + 1);
                    tlp.Controls.Add(cbCol, 3, i + 1);
                    tlp.Controls.Add(cbConv, 4, i + 1);

                    rowsUi.Add((txtName, cbDb, cbTb, cbCol, cbConv));
                }

                Button btnSave = new Button { Text = "💾 儲存設定", Dock = DockStyle.Bottom, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                btnSave.Click += (s, e) => {
                    _matConfigs.Clear();
                    foreach (var r in rowsUi) {
                        if (!string.IsNullOrWhiteSpace(r.txtName.Text) && r.cbDb.SelectedItem != null && r.cbTb.SelectedItem != null && r.cbCol.SelectedItem != null) {
                            double mult = 1.0;
                            if (r.cbConv.SelectedIndex == 1 || r.cbConv.SelectedIndex == 4) mult = 1000.0;
                            if (r.cbConv.SelectedIndex == 2 || r.cbConv.SelectedIndex == 3) mult = 0.001;

                            _matConfigs.Add(new MatConfig {
                                Alias = r.txtName.Text.Trim(),
                                DbName = r.cbDb.SelectedItem.ToString(),
                                TableName = r.cbTb.SelectedItem.ToString(),
                                ColName = r.cbCol.SelectedItem.ToString(),
                                Multiplier = mult
                            });
                        }
                    }
                    SaveMaterialConfigs();
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(tlp);
                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
        }

        private List<string> GetTablesForDb(string dbName)
        {
            List<string> result = new List<string>();
            try {
                string fullPath = Path.Combine(DataManager.BasePath, dbName + ".sqlite");
                if (File.Exists(fullPath)) {
                    using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={fullPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new System.Data.SQLite.SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'", conn))
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) result.Add(reader["name"].ToString());
                        }
                    }
                }
            } catch { }
            return result;
        }

        private async Task LoadMaterialDataAsync()
        {
            if (_matConfigs.Count == 0) {
                _dgvMaterial.DataSource = null;
                return;
            }

            int sy = int.Parse(_cboMatStartYear.SelectedItem.ToString());
            int sm = int.Parse(_cboMatSt
