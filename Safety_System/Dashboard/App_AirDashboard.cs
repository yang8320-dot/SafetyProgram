/// FILE: Safety_System/App_AirDashboard.cs ///
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
    public class App_AirDashboard
    {
        // 頂部控制項 (空污費)
        private Panel _pnlAirBox;
        private ComboBox _cboAirYear, _cboAirQuarter;
        private Label _lblAirEmissionsCurr, _lblAirEmissionsLY, _lblAirEmissionsL2Y, _lblAirEmissionsDiff;
        private Label _lblAirFeeCurr, _lblAirFeeLY, _lblAirFeeL2Y, _lblAirFeeDiff;

        // 底部控制項 (原物料統計)
        private Panel _pnlMaterialBox;
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
        
        private class AirDataResult 
        {
            public double Emissions { get; set; }
            public double Fee { get; set; }
        }

        private class MatConfigRowUI 
        {
            public TextBox txtName { get; set; }
            public ComboBox cbDb { get; set; }
            public ComboBox cbTb { get; set; }
            public ComboBox cbCol { get; set; }
            public ComboBox cbConv { get; set; }
        }

        private class ItemMap 
        {
            public string EnName { get; set; }
            public string ChName { get; set; }
            public override string ToString() => ChName;
        }

        private List<MatConfig> _matConfigs = new List<MatConfig>();
        private readonly string ConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AirMaterialSettings.txt");

        // 定義可選的資料庫清單 (提供中文選單)
        private readonly Dictionary<string, string> _dbMap = new Dictionary<string, string> {
            { "Water", "水污" }, { "Air", "空污" }, { "Waste", "廢棄物及產能" }, 
            { "Chemical", "化學品" }, { "Fire", "消防" }, { "Safety", "工安" }, { "Purchase", "請購" }
        };

        // 定義部分已知的資料表中文對照，增加設定易用性
        private readonly Dictionary<string, string> _knownTables = new Dictionary<string, string> {
            { "WaterMeterReadings", "廢水處理水量記錄" }, { "WaterChemicals", "廢水處理用藥記錄" }, { "WaterUsageDaily", "自來水使用量" }, { "DischargeData", "納管排放數據" }, { "WaterVolume", "自來水用量統計" },
            { "AirPollution", "空污申報紀錄" },
            { "WasteMonthly", "廢棄物月表" }, { "Waste_IL", "複層月表" }, { "Waste_LM", "膠合月表" }, { "Waste_CR", "鍍板月表" }, { "Waste_T", "強化月表" }, { "Waste_GCTE", "切磨月表" }, { "Waste_ML", "物料月表" }, { "Waste_Water", "水站月表" },
            { "FireResponsible", "火源責任人" }, { "HazardStats", "公共危險物統計" }, { "FireEquip", "消防設備清單" }, { "FireSelfInspection", "消防自主檢查" },
            { "SDS_Inventory", "SDS清冊" }, { "ToxicSubstances", "毒性物質" }, { "ConcernedChem", "關注性化學物質" }, { "SpecificChem", "特定化學物質" }, { "OrganicSolvents", "有機溶劑" }, { "PublicHazardous", "公共危險物品" },
            { "SafetyInspection", "巡檢記錄" }, { "WorkInjury", "工傷事件" }, { "MinorInjury", "輕傷事件" }, { "TrafficInjury", "交通意外" }, { "NearMiss", "虛驚事件" }, { "SafetyObservation", "安全觀察" },
            { "PurchaseData", "請購資料" }
        };

        public Control GetView()
        {
            LoadMaterialConfigs();

            Panel mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            Padding lblPad = new Padding(0, 8, 5, 0); 
            Padding ctrlPad = new Padding(0, 4, 10, 0); 
            int btnHeight = 35; // 統一按鈕高度，確保全排對齊

            // ==========================================
            // 第一區塊：空污費查詢與分析
            // ==========================================
            _pnlAirBox = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 30) };
            _pnlAirBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlAirBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            // 1. 標題列 (純標題，無按鈕)
            Panel pnlHeaderAir = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.White };
            Label lblAirTitle = new Label { Text = "台灣玻璃彰濱廠 - 空污費申報【排放量】統計", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DeepSkyBlue, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };
            pnlHeaderAir.Controls.Add(lblAirTitle);

            // 2. 篩選列 (採用 TableLayoutPanel 確保不崩塌)
            TableLayoutPanel tlpAirFilter = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 1, Padding = new Padding(15, 10, 15, 15) };
            tlpAirFilter.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            tlpAirFilter.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

            FlowLayoutPanel flpAirFilterLeft = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            FlowLayoutPanel flpAirFilterRight = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.RightToLeft };

            _cboAirYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100, Margin = ctrlPad };
            _cboAirQuarter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180, Margin = ctrlPad }; 
            
            int currYear = DateTime.Today.Year;
            for (int i = currYear - 10; i <= currYear; i++) _cboAirYear.Items.Add(i.ToString());
            _cboAirYear.SelectedItem = currYear.ToString();

            _cboAirQuarter.Items.AddRange(new string[] { "全年 (Q1~Q4)", "第一季", "第二季", "第三季", "第四季" });
            _cboAirQuarter.SelectedIndex = 0;

            Button btnSearchAir = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = new Padding(5, 0, 0, 0), BackColor = Color.DeepSkyBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSearchAir.FlatAppearance.BorderSize = 0;
            btnSearchAir.Click += (s, e) => LoadAirPollutionData();

            Button btnPdfAir = new Button { Text = "📄 導出 PDF", Size = new Size(130, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnPdfAir.FlatAppearance.BorderSize = 0;
            btnPdfAir.Click += (s, e) => ExportBoxToPdf(_pnlAirBox, "空污費統計報表");

            // 左側加入
            flpAirFilterLeft.Controls.AddRange(new Control[] {
                new Label { Text = "查詢年度:", AutoSize = true, Margin = lblPad }, _cboAirYear,
                new Label { Text = "申報季度:", AutoSize = true, Margin = lblPad }, _cboAirQuarter,
                btnSearchAir
            });

            // 右側加入 (RightToLeft)
            flpAirFilterRight.Controls.Add(btnPdfAir);

            tlpAirFilter.Controls.Add(flpAirFilterLeft, 0, 0);
            tlpAirFilter.Controls.Add(flpAirFilterRight, 1, 0);

            // 3. 數據方塊區
            TableLayoutPanel tlpAirData = new TableLayoutPanel { Dock = DockStyle.Top, Height = 140, ColumnCount = 4, RowCount = 2, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single, Padding = new Padding(10, 0, 10, 10) };
            for (int i = 0; i < 4; i++) tlpAirData.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpAirData.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            tlpAirData.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F)); 

            string[] airHeaders = { "當期申報數據", "去年同期數據", "前年同期數據", "與去年同期差異" };
            for (int i = 0; i < 4; i++) {
                tlpAirData.Controls.Add(new Label { Text = airHeaders[i], Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.DeepSkyBlue, ForeColor = Color.White, Margin = new Padding(0) }, i, 0);
            }

            _lblAirEmissionsCurr = CreateDataLabel(); _lblAirFeeCurr = CreateDataLabel();
            _lblAirEmissionsLY = CreateDataLabel();   _lblAirFeeLY = CreateDataLabel();
            _lblAirEmissionsL2Y = CreateDataLabel();  _lblAirFeeL2Y = CreateDataLabel();
            _lblAirEmissionsDiff = CreateDataLabel(); _lblAirFeeDiff = CreateDataLabel();

            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsCurr, _lblAirFeeCurr), 0, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsLY, _lblAirFeeLY), 1, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsL2Y, _lblAirFeeL2Y), 2, 1);
            tlpAirData.Controls.Add(CreateDataCell(_lblAirEmissionsDiff, _lblAirFeeDiff), 3, 1);

            _pnlAirBox.Controls.Add(tlpAirData);
            _pnlAirBox.Controls.Add(tlpAirFilter);
            _pnlAirBox.Controls.Add(pnlHeaderAir);

            tlpMain.Controls.Add(_pnlAirBox, 0, 0);

            // ==========================================
            // 第二區塊：原物料使用紀錄統計表
            // ==========================================
            _pnlMaterialBox = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            _pnlMaterialBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlMaterialBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            // 1. 標題列 (純標題，無按鈕)
            Panel pnlHeaderMat = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.White };
            Label lblMatTitle = new Label { Text = "台灣玻璃彰濱廠 - 原物料使用紀錄統計表", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SeaGreen, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };
            pnlHeaderMat.Controls.Add(lblMatTitle);
            
            // 2. 篩選列 (採用 TableLayoutPanel 確保不崩塌)
            TableLayoutPanel tlpMatFilter = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 1, Padding = new Padding(15, 10, 15, 15) };
            tlpMatFilter.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            tlpMatFilter.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

            FlowLayoutPanel flpMatFilterLeft = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            FlowLayoutPanel flpMatFilterRight = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.RightToLeft };
            
            _cboMatStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 90, Margin = ctrlPad };
            _cboMatStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = ctrlPad };
            _cboMatEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 90, Margin = ctrlPad };
            _cboMatEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = ctrlPad };

            for (int i = currYear - 10; i <= currYear; i++) {
                _cboMatStartYear.Items.Add(i.ToString()); _cboMatEndYear.Items.Add(i.ToString());
            }
            for (int i = 1; i <= 12; i++) {
                _cboMatStartMonth.Items.Add(i.ToString("D2")); _cboMatEndMonth.Items.Add(i.ToString("D2"));
            }

            DateTime lastYear = DateTime.Today.AddYears(-1);
            _cboMatStartYear.SelectedItem = lastYear.Year.ToString();
            _cboMatStartMonth.SelectedItem = lastYear.Month.ToString("D2");
            _cboMatEndYear.SelectedItem = DateTime.Today.Year.ToString();
            _cboMatEndMonth.SelectedItem = DateTime.Today.Month.ToString("D2");

            Button btnSearchMat = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = new Padding(5, 0, 0, 0), BackColor = Color.SeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSearchMat.FlatAppearance.BorderSize = 0;
            btnSearchMat.Click += async (s, e) => await LoadMaterialDataAsync();

            Button btnPdfMat = new Button { Text = "📄 導出 PDF", Size = new Size(120, btnHeight), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnPdfMat.FlatAppearance.BorderSize = 0;
            btnPdfMat.Click += (s, e) => ExportGridToPdf(_dgvMaterial, "原物料使用紀錄統計表");

            Button btnConfigMat = new Button { Text = "⚙️ 設定查詢", Size = new Size(130, btnHeight), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) }; 
            btnConfigMat.FlatAppearance.BorderSize = 0;
            btnConfigMat.Click += (s, e) => {
                OpenMaterialConfigDialog();
                _ = LoadMaterialDataAsync();
            };

            // 左側加入查詢
            flpMatFilterLeft.Controls.AddRange(new Control[] {
                new Label { Text = "年月區間:", AutoSize = true, Margin = lblPad }, 
                _cboMatStartYear, new Label { Text = "年", AutoSize = true, Margin = lblPad }, _cboMatStartMonth, new Label { Text = "月 ~", AutoSize = true, Margin = lblPad },
                _cboMatEndYear, new Label { Text = "年", AutoSize = true, Margin = lblPad }, _cboMatEndMonth, new Label { Text = "月", AutoSize = true, Margin = lblPad },
                btnSearchMat
            });

            // 右側加入 (RightToLeft 寫入順序)
            flpMatFilterRight.Controls.Add(btnConfigMat); // 在最右
            flpMatFilterRight.Controls.Add(btnPdfMat);    // 在次右

            tlpMatFilter.Controls.Add(flpMatFilterLeft, 0, 0);
            tlpMatFilter.Controls.Add(flpMatFilterRight, 1, 0);

            // 3. 資料表
            _dgvMaterial = new DataGridView { 
                Dock = DockStyle.Top, Height = 450, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                BorderStyle = BorderStyle.None, Margin = new Padding(10, 0, 10, 10)
            };
            _dgvMaterial.EnableHeadersVisualStyles = false;
            _dgvMaterial.ColumnHeadersDefaultCellStyle.BackColor = Color.SeaGreen;
            _dgvMaterial.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvMaterial.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvMaterial.ColumnHeadersHeight = 40;
            _dgvMaterial.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvMaterial.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 250, 245);

            _pnlMaterialBox.Controls.Add(_dgvMaterial);
            _pnlMaterialBox.Controls.Add(tlpMatFilter);
            _pnlMaterialBox.Controls.Add(pnlHeaderMat);
            
            tlpMain.Controls.Add(_pnlMaterialBox, 0, 1);

            mainPanel.Controls.Add(tlpMain);

            LoadAirPollutionData();
            _ = LoadMaterialDataAsync();

            return mainPanel;
        }

        // ====================================================
        // Part 1 & 2: 空污費邏輯
        // ====================================================
        private Label CreateDataLabel() => new Label { Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.FromArgb(50, 50, 50), AutoSize = true, Margin = new Padding(0, 5, 0, 10) };

        private FlowLayoutPanel CreateDataCell(Label l1, Label l2)
        {
            FlowLayoutPanel flp = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                FlowDirection = FlowDirection.TopDown, 
                AutoSize = true, 
                WrapContents = false, 
                Padding = new Padding(15, 10, 15, 10) 
            };
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

            _lblAirEmissionsDiff.Text = $"排放總量：{CalculateDiff(currData.Emissions, lyData.Emissions)}";
            _lblAirEmissionsDiff.ForeColor = currData.Emissions > lyData.Emissions ? Color.IndianRed : Color.ForestGreen;
            
            _lblAirFeeDiff.Text = $"繳費金額：{CalculateDiff(currData.Fee, lyData.Fee)}";
            _lblAirFeeDiff.ForeColor = currData.Fee > lyData.Fee ? Color.IndianRed : Color.ForestGreen;
        }

        private AirDataResult FetchAirData(int year, string quarterMode)
        {
            var res = new AirDataResult();
            try {
                DataTable dt = DataManager.GetTableData("Air", "AirPollution", "", "", "");
                if (dt != null) {
                    foreach (DataRow r in dt.Rows) {
                        string dbYear = r["年度"]?.ToString().Trim();
                        string dbQtr = r["季度"]?.ToString().Trim();

                        if (dbYear == year.ToString()) {
                            bool matchQtr = false;
                            if (quarterMode.Contains("全年")) matchQtr = true;
                            else if (quarterMode == "第一季" && (dbQtr.Contains("1") || dbQtr.Contains("一") || dbQtr.ToUpper().Contains("Q1"))) matchQtr = true;
                            else if (quarterMode == "第二季" && (dbQtr.Contains("2") || dbQtr.Contains("二") || dbQtr.ToUpper().Contains("Q2"))) matchQtr = true;
                            else if (quarterMode == "第三季" && (dbQtr.Contains("3") || dbQtr.Contains("三") || dbQtr.ToUpper().Contains("Q3"))) matchQtr = true;
                            else if (quarterMode == "第四季" && (dbQtr.Contains("4") || dbQtr.Contains("四") || dbQtr.ToUpper().Contains("Q4"))) matchQtr = true;

                            if (matchQtr) {
                                if (double.TryParse(r["排放量"]?.ToString().Replace(",", ""), out double em)) res.Emissions += em;
                                if (double.TryParse(r["繳費金額"]?.ToString().Replace(",", ""), out double f)) res.Fee += f;
                            }
                        }
                    }
                }
            } catch { }
            return res;
        }

        private void UpdateAirLabels(Label lEmissions, Label lFee, AirDataResult data)
        {
            lEmissions.Text = $"排放總量：{data.Emissions:N2} kg";
            lFee.Text = $"繳費金額：{data.Fee:N0} NTD";
        }

        private string CalculateDiff(double curr, double ly)
        {
            if (ly == 0) return "無基期";
            double diff = ((curr - ly) / ly) * 100;
            return (diff > 0 ? "+" : "") + diff.ToString("N1") + " %";
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
            using (Form f = new Form { Text = "⚙️ 設定原物料查詢組合 (最多 10 組)", Size = new Size(1100, 650), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Panel pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 5, RowCount = 11, Padding = new Padding(15) };
                
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F)); // 名稱
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F)); // 庫 (加寬給中文)
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F)); // 表
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F)); // 欄位
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));  // 換算

                string[] headers = { "顯示名稱 (如: 柴油(公秉))", "來源資料庫", "來源資料表", "加總欄位", "單位換算規則" };
                for (int i = 0; i < 5; i++) tlp.Controls.Add(new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) }, i, 0);

                var rowsUi = new List<MatConfigRowUI>();

                for (int i = 0; i < 10; i++)
                {
                    TextBox txtName = new TextBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbTb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                    ComboBox cbConv = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

                    foreach (var kvp in _dbMap) {
                        cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value });
                    }

                    cbConv.Items.AddRange(new string[] { "無換算 (x1)", "公噸 ➔ 公斤 (x1000)", "公斤 ➔ 公噸 (x0.001)", "公升 ➔ 公秉 (x0.001)", "公秉 ➔ 公升 (x1000)" });

                    cbDb.SelectedIndexChanged += (s, e) => {
                        cbTb.Items.Clear(); cbCol.Items.Clear();
                        if (cbDb.SelectedItem != null) {
                            string enDb = ((ItemMap)cbDb.SelectedItem).EnName;
                            var tbs = GetTablesForDb(enDb);
                            foreach(var tb in tbs) {
                                string chTb = _knownTables.ContainsKey(tb) ? _knownTables[tb] : tb;
                                cbTb.Items.Add(new ItemMap { EnName = tb, ChName = chTb });
                            }
                        }
                    };

                    cbTb.SelectedIndexChanged += (s, e) => {
                        cbCol.Items.Clear();
                        if (cbDb.SelectedItem != null && cbTb.SelectedItem != null) {
                            string enDb = ((ItemMap)cbDb.SelectedItem).EnName;
                            string enTb = ((ItemMap)cbTb.SelectedItem).EnName;
                            var cols = DataManager.GetColumnNames(enDb, enTb);
                            foreach(var c in cols) {
                                if (c != "Id" && c != "日期" && c != "年月" && c != "年度" && c != "備註" && c != "附件檔案") {
                                    cbCol.Items.Add(c);
                                }
                            }
                        }
                    };

                    if (i < _matConfigs.Count) {
                        var conf = _matConfigs[i];
                        txtName.Text = conf.Alias;
                        
                        foreach (ItemMap item in cbDb.Items) {
                            if (item.EnName == conf.DbName) { cbDb.SelectedItem = item; break; }
                        }

                        if (cbTb.Items.Count > 0) {
                            foreach (ItemMap item in cbTb.Items) {
                                if (item.EnName == conf.TableName) { cbTb.SelectedItem = item; break; }
                            }
                        }

                        if (cbCol.Items.Contains(conf.ColName)) cbCol.SelectedItem = conf.ColName;
                        
                        if (conf.Multiplier == 1000) cbConv.SelectedIndex = 1; 
                        else if (conf.Multiplier == 0.001) cbConv.SelectedIndex = 2; 
                        else cbConv.SelectedIndex = 0;
                    } else {
                        cbConv.SelectedIndex = 0;
                    }

                    tlp.Controls.Add(txtName, 0, i + 1);
                    tlp.Controls.Add(cbDb, 1, i + 1);
                    tlp.Controls.Add(cbTb, 2, i + 1);
                    tlp.Controls.Add(cbCol, 3, i + 1);
                    tlp.Controls.Add(cbConv, 4, i + 1);

                    rowsUi.Add(new MatConfigRowUI { txtName = txtName, cbDb = cbDb, cbTb = cbTb, cbCol = cbCol, cbConv = cbConv });
                }

                pnlScroll.Controls.Add(tlp);

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
                                DbName = ((ItemMap)r.cbDb.SelectedItem).EnName,
                                TableName = ((ItemMap)r.cbTb.SelectedItem).EnName,
                                ColName = r.cbCol.SelectedItem.ToString(),
                                Multiplier = mult
                            });
                        }
                    }
                    SaveMaterialConfigs();
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(pnlScroll);
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
            int sm = int.Parse(_cboMatStartMonth.SelectedItem.ToString());
            int ey = int.Parse(_cboMatEndYear.SelectedItem.ToString());
            int em = int.Parse(_cboMatEndMonth.SelectedItem.ToString());

            DateTime start = new DateTime(sy, sm, 1);
            DateTime end = new DateTime(ey, em, 1);
            if (start > end) { MessageBox.Show("起始年月不能大於結束年月！"); return; }

            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("年月", typeof(string));
            foreach (var conf in _matConfigs) dtResult.Columns.Add(conf.Alias, typeof(double));

            double[] totals = new double[_matConfigs.Count];

            await Task.Run(() => {
                DateTime curr = start;
                while (curr <= end) {
                    string ymStr = curr.ToString("yyyy-MM");
                    DataRow row = dtResult.NewRow();
                    row["年月"] = ymStr;

                    for (int i = 0; i < _matConfigs.Count; i++) {
                        var conf = _matConfigs[i];
                        double sum = GetMonthlySum(conf.DbName, conf.TableName, conf.ColName, ymStr);
                        sum *= conf.Multiplier;
                        row[conf.Alias] = sum;
                        totals[i] += sum;
                    }
                    dtResult.Rows.Add(row);
                    curr = curr.AddMonths(1);
                }
            });

            DataRow sumRow = dtResult.NewRow();
            sumRow["年月"] = "【合計】";
            for (int i = 0; i < totals.Length; i++) sumRow[_matConfigs[i].Alias] = totals[i];
            dtResult.Rows.Add(sumRow);

            _dgvMaterial.DataSource = dtResult;

            if (_dgvMaterial.Columns.Contains("年月")) {
                _dgvMaterial.Columns["年月"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                _dgvMaterial.Columns["年月"].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            }

            foreach (var conf in _matConfigs) {
                if (_dgvMaterial.Columns.Contains(conf.Alias)) {
                    _dgvMaterial.Columns[conf.Alias].DefaultCellStyle.Format = "N2";
                    _dgvMaterial.Columns[conf.Alias].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; 
                }
            }

            if (_dgvMaterial.Rows.Count > 0) {
                var lastRow = _dgvMaterial.Rows[_dgvMaterial.Rows.Count - 1];
                lastRow.DefaultCellStyle.BackColor = Color.LightYellow;
                lastRow.DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                lastRow.DefaultCellStyle.ForeColor = Color.DarkRed;
            }

            _dgvMaterial.ClearSelection();
        }

        private double GetMonthlySum(string db, string table, string col, string ymStr)
        {
            double total = 0;
            try {
                DataTable dt = DataManager.GetTableData(db, table, "", "", ""); 
                if (dt != null && dt.Columns.Contains(col)) {
                    string dateCol = dt.Columns.Contains("年月") ? "年月" : (dt.Columns.Contains("日期") ? "日期" : "");
                    if (string.IsNullOrEmpty(dateCol)) return 0;

                    foreach (DataRow r in dt.Rows) {
                        string dVal = r[dateCol]?.ToString() ?? "";
                        if (dVal.StartsWith(ymStr)) {
                            if (double.TryParse(r[col]?.ToString().Replace(",", ""), out double v)) {
                                total += v;
                            }
                        }
                    }
                }
            } catch { }
            return total;
        }

        // ====================================================
        // PDF 高清向量導出系統
        // ====================================================
        private void ExportBoxToPdf(Panel pnlBox, string title)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = true;
                        
                        pd.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(30, 30, 40, 40);

                        pd.PrintPage += (s, ev) => {
                            Graphics g = ev.Graphics;
                            float x = ev.MarginBounds.Left; 
                            float y = ev.MarginBounds.Top; 
                            float pageWidth = ev.MarginBounds.Width;

                            Font fMainTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold); 
                            Font fSubTitle = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold); 
                            Font fSign = new Font("Microsoft JhengHei UI", 10F); 
                            Font fDate = new Font("Microsoft JhengHei UI", 10F); 
                            
                            Font fBody = new Font("Microsoft JhengHei UI", 12F); 
                            Font fHead = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                            
                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }; 
                            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                            g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fMainTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 30), sfCenter); 
                            y += 35;
                            g.DrawString("空污費申報【排放量】統計報表", fSubTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 25), sfCenter); 
                            y += 45; 
                            
                            string signText = "廠主管：_________________        經/副理：_________________        課/股長：_________________        主辦：_________________        製表人：_________________";
                            g.DrawString(signText, fSign, Brushes.Black, new RectangleF(x, y, pageWidth, 20), sfCenter);
                            y += 45; 

                            g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}    查詢區間：{_cboAirYear.Text}年度 {_cboAirQuarter.Text}", fDate, Brushes.Gray, x, y); 
                            y += 25;
                            
                            string[] headers = { "當期申報數據", "去年同期數據", "前年同期數據", "與去年同期差異" };
                            string[] emissions = { _lblAirEmissionsCurr.Text, _lblAirEmissionsLY.Text, _lblAirEmissionsL2Y.Text, _lblAirEmissionsDiff.Text };
                            string[] fees = { _lblAirFeeCurr.Text, _lblAirFeeLY.Text, _lblAirFeeL2Y.Text, _lblAirFeeDiff.Text };

                            float colWidth = pageWidth / 4;
                            float headerH = 40;
                            float rowH = 80;

                            for(int i=0; i<4; i++) {
                                RectangleF rectF = new RectangleF(x + i*colWidth, y, colWidth, headerH);
                                Rectangle rect = Rectangle.Round(rectF);
                                g.FillRectangle(Brushes.DeepSkyBlue, rectF);
                                g.DrawRectangle(Pens.Black, rect);
                                g.DrawString(headers[i], fHead, Brushes.White, rectF, sfCenter);
                            }
                            y += headerH;

                            for(int i=0; i<4; i++) {
                                RectangleF rectF = new RectangleF(x + i*colWidth, y, colWidth, rowH);
                                Rectangle rect = Rectangle.Round(rectF);
                                g.DrawRectangle(Pens.Black, rect);
                                string text = $"{emissions[i]}\n\n{fees[i]}";
                                g.DrawString(text, fBody, Brushes.Black, new RectangleF(rectF.X + 10, rectF.Y + 10, rectF.Width - 20, rectF.Height - 20), sfLeft);
                            }

                            g.DrawString($"第 1 頁 / 共 1 頁", fDate, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);

                            ev.HasMorePages = false;
                        };

                        pd.Print();
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void ExportGridToPdf(DataGridView dgv, string title)
        {
            if (dgv.Rows.Count == 0) { MessageBox.Show("目前沒有資料可供導出。"); return; }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    Form activeForm = Form.ActiveForm; 
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;
                    
                    PrintDocument pd = new PrintDocument(); 
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF"; 
                    pd.PrinterSettings.PrintToFile = true; 
                    pd.PrinterSettings.PrintFileName = sfd.FileName; 
                    pd.DefaultPageSettings.Landscape = true; 
                    
                    pd.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(30, 30, 40, 40);
                    
                    int rowIndex = 0; 
                    int pageNumber = 1; 

                    int totalPages = 1;
                    if (dgv.Rows.Count > 15) {
                        totalPages = 1 + (int)Math.Ceiling((double)(dgv.Rows.Count - 15) / 18);
                    }
                    
                    pd.PrintPage += (s, ev) => 
                    {
                        Graphics g = ev.Graphics; 
                        float x = ev.MarginBounds.Left; 
                        float y = ev.MarginBounds.Top; 
                        float pageWidth = ev.MarginBounds.Width;
                        
                        Font fMainTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold); 
                        Font fSubTitle = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold); 
                        Font fSign = new Font("Microsoft JhengHei UI", 10F); 
                        Font fDate = new Font("Microsoft JhengHei UI", 10F); 

                        Font fBody = new Font("Microsoft JhengHei UI", 10F); 
                        Font fHead = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
                        
                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }; 
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        if (rowIndex == 0) {
                            g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fMainTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 30), sfCenter); 
                            y += 35;
                            g.DrawString(title, fSubTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 25), sfCenter); 
                            y += 45; 

                            string signText = "廠主管：_________________        經/副理：_________________        課/股長：_________________        主辦：_________________        製表人：_________________";
                            g.DrawString(signText, fSign, Brushes.Black, new RectangleF(x, y, pageWidth, 20), sfCenter);
                            y += 45; 

                            g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}", fDate, Brushes.Gray, x, y); 
                            y += 25;
                        } else {
                            g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}", fDate, Brushes.Gray, x, y); 
                            y += 25;
                        }
                        
                        var visCols = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList(); 
                        if (visCols.Count == 0) return;
                        
                        float totalGridWidth = visCols.Sum(c => c.Width); 
                        float[] colWidths = new float[visCols.Count];
                        for (int i = 0; i < visCols.Count; i++) colWidths[i] = (visCols[i].Width / totalGridWidth) * pageWidth; 
                        
                        float currX = x; 
                        float rowH = 35;
                        
                        for (int i = 0; i < visCols.Count; i++) 
                        {
                            RectangleF rectF = new RectangleF(currX, y, colWidths[i], rowH);
                            Rectangle rect = Rectangle.Round(rectF);
                            g.FillRectangle(Brushes.SeaGreen, rectF); 
                            g.DrawRectangle(Pens.Black, rect); 
                            g.DrawString(visCols[i].HeaderText, fHead, Brushes.White, rectF, sfCenter); 
                            currX += colWidths[i];
                        }
                        y += rowH;
                        
                        while (rowIndex < dgv.Rows.Count) 
                        {
                            if (y + rowH > ev.MarginBounds.Bottom - 30) 
                            { 
                                g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fDate, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter); 
                                pageNumber++; 
                                ev.HasMorePages = true; 
                                return; 
                            }
                            
                            currX = x;
                            for (int i = 0; i < visCols.Count; i++) 
                            {
                                RectangleF rectF = new RectangleF(currX, y, colWidths[i], rowH); 
                                Rectangle rect = Rectangle.Round(rectF);
                                g.DrawRectangle(Pens.Black, rect);
                                string val = dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? ""; 
                                g.DrawString(val, fBody, Brushes.Black, rectF, sfCenter); 
                                currX += colWidths[i];
                            }
                            y += rowH; 
                            rowIndex++;
                        }
                        
                        g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fDate, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter); 
                        ev.HasMorePages = false; 
                        rowIndex = 0; 
                        pageNumber = 1;
                    };
                    
                    try 
                    { 
                        pd.Print(); 
                        if (activeForm != null) activeForm.Cursor = Cursors.Default; 
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    } 
                    catch (Exception ex) 
                    { 
                        if (activeForm != null) activeForm.Cursor = Cursors.Default; 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } 
                }
            }
        }
    }
}
