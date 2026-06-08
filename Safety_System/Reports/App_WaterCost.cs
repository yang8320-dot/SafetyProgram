using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_WaterCost
    {
        // 頂部日期區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        private Button _btnSearch;

        private class DashboardSectionUI {
            public Panel MainBox;
            public Label LblSub1, LblSub2, LblSub3, LblSub4;
            public FlowLayoutPanel PnlData1, PnlData2, PnlData3, PnlData4;
            public Label LblTotalExpense; 
            public Label LblTotalBenefit; 
            public Label LblTotalCarbon; 
        }

        private DashboardSectionUI _sec1, _sec2, _sec3, _sec4, _sec5;

        // 資料庫與快取
        private const string SysDbName = "SystemConfig";
        private const string ConfigTable = "WaterCostFormulas";
        private const string PriceTable = "WaterPrices";
        private const string PriceCalcTemplateTable = "WaterPriceCalcTemplates"; // 🟢 新增：公式範本資料表

        private List<CostFormulaItem> _configs = new List<CostFormulaItem>();
        private List<PriceItem> _prices = new List<PriceItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 模型定義
        private class CostFormulaItem {
            public int Id { get; set; }
            public string Section { get; set; }     
            public string DisplayName { get; set; }
            public string OutputType { get; set; }  
            public string CostType { get; set; } 
            public string Unit { get; set; } 
            public string Formula { get; set; }     
        }

        private class PriceItem {
            public int Id { get; set; }
            public string Category { get; set; } 
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            public double UnitPrice { get; set; }
        }

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public Control GetView()
        {
            _dbMap = App_DbConfig.GetDbMapCache();
            InitDatabase();
            LoadCache();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            
            TableLayoutPanel masterLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 7, Margin = new Padding(0) };
            for(int i=0; i<7; i++) masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ==========================================
            // 1. 標題與操作區
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "🌱 水資源成本與 ESG 效益分析看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblTitle);

            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 10, 0, 20), Margin = new Padding(0) };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes();

            _btnSearch = new Button { Text = "🔍 執行精算", Size = new Size(130, 42), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.Click += (s, e) => ExecuteCalculation();

            Button btnPriceManager = new Button { Text = "💰 費率與碳排係數管理", Size = new Size(230, 42), BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnPriceManager.Click += (s, e) => { OpenPriceManager(); ExecuteCalculation(); };

            // 🟢 將統計設定按鈕統一移到此處
            Button btnUnifiedSetting = new Button { Text = "⚙️ 公式與統計設定", Size = new Size(200, 42), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnUnifiedSetting.Click += (s, e) => { OpenUnifiedConfigManager(); ExecuteCalculation(); };

            Button btnPdf = new Button { Text = "📄 選擇並導出 PDF", Size = new Size(180, 42), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnPdf.Click += (s,e) => ExportToPdf();

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "計算區間:", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 10, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 10, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 10, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _btnSearch, btnPriceManager, btnUnifiedSetting, btnPdf
            });

            // ==========================================
            // 2. 建立五大區塊
            // ==========================================
            _sec1 = BuildSection("廢水處理費用統計", "廢水處理", Color.Sienna);
            _sec2 = BuildSection("淨水處理費用統計", "淨水處理", Color.MediumBlue);
            _sec3 = BuildSection("回收水成本與效益", "回收水", Color.ForestGreen);
            _sec4 = BuildSection("雨水回收效益分析", "雨水回收", Color.SteelBlue); 
            _sec5 = BuildSection("污泥減量與處置效益", "污泥減量", Color.Purple); 

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);
            masterLayout.Controls.Add(_sec1.MainBox, 0, 2);
            masterLayout.Controls.Add(_sec2.MainBox, 0, 3);
            masterLayout.Controls.Add(_sec3.MainBox, 0, 4);
            masterLayout.Controls.Add(_sec4.MainBox, 0, 5); 
            masterLayout.Controls.Add(_sec5.MainBox, 0, 6); 

            mainScrollPanel.Controls.Add(masterLayout);

            ExecuteCalculation();

            return mainScrollPanel;
        }

        // ==========================================
        // 資料庫初始化與快取
        // ==========================================
        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql1 = $"CREATE TABLE IF NOT EXISTS [{ConfigTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, Section TEXT, DisplayName TEXT, OutputType TEXT, Unit TEXT, Formula TEXT);";
                    string sql2 = $"CREATE TABLE IF NOT EXISTS [{PriceTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, Category TEXT, StartDate TEXT, EndDate TEXT, UnitPrice REAL);";
                    string sql3 = $"CREATE TABLE IF NOT EXISTS [{PriceCalcTemplateTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TemplateName TEXT, Formula TEXT);"; // 🟢 建立範本資料表
                    
                    using (var cmd = new SQLiteCommand(sql1, conn)) { cmd.ExecuteNonQuery(); }
                    using (var cmd = new SQLiteCommand(sql2, conn)) { cmd.ExecuteNonQuery(); }
                    using (var cmd = new SQLiteCommand(sql3, conn)) { cmd.ExecuteNonQuery(); }

                    var cols = new List<string>();
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info([{ConfigTable}])", conn))
                    using (var r = cmd.ExecuteReader()) {
                        while (r.Read()) cols.Add(r["name"].ToString());
                    }
                    if (!cols.Contains("Unit")) {
                        using (var cmd = new SQLiteCommand($"ALTER TABLE [{ConfigTable}] ADD COLUMN Unit TEXT;", conn)) { cmd.ExecuteNonQuery(); }
                    }
                    if (!cols.Contains("CostType")) {
                        using (var cmd = new SQLiteCommand($"ALTER TABLE [{ConfigTable}] ADD COLUMN CostType TEXT;", conn)) { cmd.ExecuteNonQuery(); }
                    }
                }
            } catch (Exception ex) { 
                MessageBox.Show("初始化設定資料表發生異常：" + ex.Message);
            }
        }

        private void LoadCache()
        {
            _configs.Clear();
            _prices.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"SELECT * FROM {ConfigTable}", conn))
                    using (var r = cmd.ExecuteReader()) {
                        bool hasUnitCol = false;
                        bool hasCostTypeCol = false;
                        for (int i = 0; i < r.FieldCount; i++) {
                            if (r.GetName(i).Equals("Unit", StringComparison.OrdinalIgnoreCase)) hasUnitCol = true;
                            if (r.GetName(i).Equals("CostType", StringComparison.OrdinalIgnoreCase)) hasCostTypeCol = true;
                        }

                        while (r.Read()) {
                            string unit = hasUnitCol ? r["Unit"].ToString() : "";
                            string costType = hasCostTypeCol ? r["CostType"].ToString() : "無";
                            _configs.Add(new CostFormulaItem { 
                                Id = Convert.ToInt32(r["Id"]), Section = r["Section"].ToString(), DisplayName = r["DisplayName"].ToString(), 
                                OutputType = r["OutputType"].ToString(), CostType = costType, Unit = unit, Formula = r["Formula"].ToString()
                            });
                        }
                    }

                    using (var cmd = new SQLiteCommand($"SELECT * FROM {PriceTable}", conn))
                    using (var r = cmd.ExecuteReader()) {
                        while (r.Read()) _prices.Add(new PriceItem { 
                            Id = Convert.ToInt32(r["Id"]), Category = r["Category"].ToString(), 
                            StartDate = DateTime.Parse(r["StartDate"].ToString()), EndDate = DateTime.Parse(r["EndDate"].ToString()), 
                            UnitPrice = Convert.ToDouble(r["UnitPrice"]) 
                        });
                    }
                }
            } catch { }
        }

        // ==========================================
        // UI 區塊建立 
        // ==========================================
        private DashboardSectionUI BuildSection(string title, string sectionCode, Color themeColor)
        {
            DashboardSectionUI ui = new DashboardSectionUI();

            ui.MainBox = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 25) };
            ui.MainBox.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, ui.MainBox.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.White };
            Label lblTitle = new Label { Text = $"■ {title}", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = themeColor, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, AutoSize = true, Padding=new Padding(10,15,0,0) };
            
            FlowLayoutPanel flpRight = new FlowLayoutPanel { 
                Dock = DockStyle.Right, 
                AutoSize = true, 
                FlowDirection = FlowDirection.RightToLeft, 
                WrapContents = false,
                Padding = new Padding(0, 10, 15, 0)
            };

            // 🟢 移除了各區塊右上角的獨立「設定按鈕」
            ui.LblTotalBenefit = new Label { Text = "總效益: $ 0", Font = new Font("Consolas", 15F, FontStyle.Bold), ForeColor = Color.SteelBlue, TextAlign = ContentAlignment.MiddleRight, AutoSize = true, Margin = new Padding(15, 8, 0, 0) };
            ui.LblTotalExpense = new Label { Text = "總支出: $ 0", Font = new Font("Consolas", 15F, FontStyle.Bold), ForeColor = Color.Crimson, TextAlign = ContentAlignment.MiddleRight, AutoSize = true, Margin = new Padding(15, 8, 0, 0) };
            ui.LblTotalCarbon = new Label { Text = "總碳排當量: 0 kgCO2e", Font = new Font("Consolas", 15F, FontStyle.Bold), ForeColor = Color.DarkOliveGreen, TextAlign = ContentAlignment.MiddleRight, AutoSize = true, Margin = new Padding(15, 8, 0, 0) };

            flpRight.Controls.Add(ui.LblTotalBenefit);
            flpRight.Controls.Add(ui.LblTotalExpense);
            flpRight.Controls.Add(ui.LblTotalCarbon);

            pnlHeader.Controls.Add(flpRight);
            pnlHeader.Controls.Add(lblTitle);

            TableLayoutPanel gridFour = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, RowCount = 2, Padding = new Padding(10) };
            for (int i = 0; i < 4; i++) gridFour.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));
            gridFour.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            ui.LblSub1 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            ui.LblSub2 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            ui.LblSub3 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            ui.LblSub4 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = themeColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };

            gridFour.Controls.Add(ui.LblSub1, 0, 0); gridFour.Controls.Add(ui.LblSub2, 1, 0);
            gridFour.Controls.Add(ui.LblSub3, 2, 0); gridFour.Controls.Add(ui.LblSub4, 3, 0);

            ui.PnlData1 = CreateDataPanel(); ui.PnlData2 = CreateDataPanel();
            ui.PnlData3 = CreateDataPanel(); ui.PnlData4 = CreateDataPanel();

            gridFour.Controls.Add(ui.PnlData1, 0, 1); gridFour.Controls.Add(ui.PnlData2, 1, 1);
            gridFour.Controls.Add(ui.PnlData3, 2, 1); gridFour.Controls.Add(ui.PnlData4, 3, 1);

            ui.MainBox.Controls.Add(gridFour);
            ui.MainBox.Controls.Add(pnlHeader);

            return ui;
        }

        private FlowLayoutPanel CreateDataPanel()
        {
            return new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 100), FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.FromArgb(248, 249, 250), Margin = new Padding(2), Padding = new Padding(10) };
        }

        // ==========================================
        // 核心計算引擎
        // ==========================================
        private void ExecuteCalculation()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            RenderSectionData("廢水處理", _sec1, dtS, dtE, Color.Sienna);
            RenderSectionData("淨水處理", _sec2, dtS, dtE, Color.MediumBlue);
            RenderSectionData("回收水", _sec3, dtS, dtE, Color.ForestGreen);
            RenderSectionData("雨水回收", _sec4, dtS, dtE, Color.SteelBlue); 
            RenderSectionData("污泥減量", _sec5, dtS, dtE, Color.Purple); 

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private void RenderSectionData(string sectionCode, DashboardSectionUI ui, DateTime dtS, DateTime dtE, Color themeColor)
        {
            ui.PnlData1.Controls.Clear(); ui.PnlData2.Controls.Clear(); 
            ui.PnlData3.Controls.Clear(); ui.PnlData4.Controls.Clear();

            var sectionConfigs = _configs.Where(c => c.Section == sectionCode).ToList();

            if (sectionConfigs.Count == 0) {
                ui.PnlData1.Controls.Add(new Label { Text = "尚未設定任何統計項目，請點擊上方 [公式與統計設定] 進行新增。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                ui.LblTotalExpense.Text = "總支出: $ 0";
                ui.LblTotalBenefit.Text = "總效益: $ 0";
                ui.LblTotalCarbon.Text = "總碳排當量: 0 kgCO2e";
                return;
            }

            ui.LblSub1.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n區間統計總計";
            ui.LblSub2.Text = $"【{dtS.AddYears(-1):yyyy/MM/dd} ~ {dtE.AddYears(-1):yyyy/MM/dd}】\n去年同期統計總計";
            ui.LblSub3.Text = $"【{dtS.AddYears(-2):yyyy/MM/dd} ~ {dtE.AddYears(-2):yyyy/MM/dd}】\n前年同期統計總計";
            ui.LblSub4.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n與去年同期差異分析";

            double sectionTotalExpense = 0;
            double sectionTotalBenefit = 0;
            double sectionTotalCarbon = 0; 

            foreach (var cfg in sectionConfigs)
            {
                double vCurr = EvaluateFormula(cfg.Formula, dtS, dtE);
                double vLy   = EvaluateFormula(cfg.Formula, dtS.AddYears(-1), dtE.AddYears(-1));
                double vL2y  = EvaluateFormula(cfg.Formula, dtS.AddYears(-2), dtE.AddYears(-2));

                string unit = string.IsNullOrEmpty(cfg.Unit) ? "" : cfg.Unit;
                string prefix = "";

                if (cfg.OutputType == "金額") { 
                    if (string.IsNullOrEmpty(unit)) unit = "元"; 
                    prefix = "$"; 
                }
                else if (cfg.OutputType == "碳排(kgCO2e)") { 
                    if (string.IsNullOrEmpty(unit)) unit = "kgCO2e"; 
                    prefix = "☁️"; 
                }

                ui.PnlData1.Controls.Add(CreateStatLabel(cfg.DisplayName, vCurr, unit, prefix, themeColor));
                ui.PnlData2.Controls.Add(CreateStatLabel(cfg.DisplayName, vLy, unit, prefix, themeColor));
                ui.PnlData3.Controls.Add(CreateStatLabel(cfg.DisplayName, vL2y, unit, prefix, themeColor));

                double diff = vCurr - vLy;
                double yoy = vLy == 0 ? 0 : (diff / Math.Abs(vLy)) * 100;

                string diffText = vLy == 0 && vCurr > 0 ? "新數據" : (vLy == 0 ? "無基期" : $"{(yoy > 0 ? "+" : "")}{yoy:N1} %");
                Color diffColor = (vLy == 0 && vCurr > 0) ? Color.SteelBlue : (vLy == 0 ? Color.DimGray : (yoy > 0 ? Color.IndianRed : Color.ForestGreen));

                ui.PnlData4.Controls.Add(new Label { 
                    Text = $"{cfg.DisplayName}: {diffText}", 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = diffColor, 
                    AutoSize = true, 
                    Margin = new Padding(0, 0, 0, 8) 
                });

                if (cfg.OutputType == "金額") {
                    if (cfg.CostType == "支出") sectionTotalExpense += vCurr;
                    else if (cfg.CostType == "效益") sectionTotalBenefit += vCurr;
                }
                else if (cfg.OutputType == "碳排(kgCO2e)") {
                    sectionTotalCarbon += vCurr;
                }
            }

            ui.LblTotalExpense.Text = $"總支出: $ {sectionTotalExpense:N0}";
            ui.LblTotalBenefit.Text = $"總效益: $ {sectionTotalBenefit:N0}";
            ui.LblTotalCarbon.Text = $"總碳排當量: {sectionTotalCarbon:N1} kgCO2e";
        }

        private Label CreateStatLabel(string title, double value, string unit, string prefix, Color themeColor)
        {
            string valStr = prefix == "$" ? $"{value:N0}" : $"{value:N2}";
            string fullText = $"{title}: {prefix} {valStr} {unit}".Trim();
            
            return new Label { 
                Text = fullText, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                ForeColor = Color.FromArgb(60, 60, 60), 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 10) 
            };
        }

        private double EvaluateFormula(string formula, DateTime dtS, DateTime dtE)
        {
            string sStr = dtS.ToString("yyyy-MM-dd");
            string eStr = dtE.ToString("yyyy-MM-dd");
            string sYm = dtS.ToString("yyyy-MM");
            string eYm = dtE.ToString("yyyy-MM");

            // 1. COST()
            Regex costRegex = new Regex(@"COST\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\],\s*(?<cat>[^\)]+)\)");
            var costMatches = costRegex.Matches(formula);
            foreach (Match m in costMatches) {
                string db = m.Groups["db"].Value; string tb = m.Groups["tb"].Value;
                string col = m.Groups["col"].Value; string cat = m.Groups["cat"].Value.Trim();
                double costSum = 0;
                try {
                    var allCols = DataManager.GetColumnNames(db, tb);
                    string dateCol = allCols.Contains("日期") ? "日期" : (allCols.Contains("年月") ? "年月" : "");
                    if (!string.IsNullOrEmpty(dateCol)) {
                        string qS = dateCol == "年月" ? sYm : sStr;
                        string qE = dateCol == "年月" ? eYm : eStr;
                        DataTable dt = DataManager.GetTableData(db, tb, dateCol, qS, qE);
                        if (dt != null && dt.Columns.Contains(col)) {
                            foreach (DataRow r in dt.Rows) {
                                if (double.TryParse(r[col]?.ToString().Replace(",",""), out double qty)) {
                                    DateTime rowDate = dtS; 
                                    if (dateCol == "日期") DateTime.TryParse(r["日期"].ToString(), out rowDate);
                                    else if (dateCol == "年月") DateTime.TryParse(r["年月"].ToString() + "-01", out rowDate);
                                    double price = GetPriceForDate(cat, rowDate);
                                    costSum += (qty * price);
                                }
                            }
                        }
                    }
                } catch { }
                formula = formula.Replace(m.Value, costSum.ToString());
            }

            // 2. SUM()
            Regex sumRegex = new Regex(@"SUM\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
            var sumMatches = sumRegex.Matches(formula);
            foreach (Match m in sumMatches) {
                string db = m.Groups["db"].Value; string tb = m.Groups["tb"].Value; string col = m.Groups["col"].Value;
                double qtySum = 0;
                try {
                    var allCols = DataManager.GetColumnNames(db, tb);
                    string dateCol = allCols.Contains("日期") ? "日期" : (allCols.Contains("年月") ? "年月" : "");
                    if (!string.IsNullOrEmpty(dateCol)) {
                        string qS = dateCol == "年月" ? sYm : sStr;
                        string qE = dateCol == "年月" ? eYm : eStr;
                        DataTable dt = DataManager.GetTableData(db, tb, dateCol, qS, qE);
                        if (dt != null && dt.Columns.Contains(col)) {
                            foreach (DataRow r in dt.Rows) {
                                if (double.TryParse(r[col]?.ToString().Replace(",",""), out double qty)) qtySum += qty;
                            }
                        }
                    }
                } catch { }
                formula = formula.Replace(m.Value, qtySum.ToString());
            }

            // 3. AVG()
            Regex avgRegex = new Regex(@"AVG\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
            var avgMatches = avgRegex.Matches(formula);
            foreach (Match m in avgMatches) {
                string db = m.Groups["db"].Value; string tb = m.Groups["tb"].Value; string col = m.Groups["col"].Value;
                double qtySum = 0; int qtyCount = 0;
                try {
                    var allCols = DataManager.GetColumnNames(db, tb);
                    string dateCol = allCols.Contains("日期") ? "日期" : (allCols.Contains("年月") ? "年月" : "");
                    if (!string.IsNullOrEmpty(dateCol)) {
                        string qS = dateCol == "年月" ? sYm : sStr;
                        string qE = dateCol == "年月" ? eYm : eStr;
                        DataTable dt = DataManager.GetTableData(db, tb, dateCol, qS, qE);
                        if (dt != null && dt.Columns.Contains(col)) {
                            foreach (DataRow r in dt.Rows) {
                                if (double.TryParse(r[col]?.ToString().Replace(",",""), out double qty)) {
                                    qtySum += qty; qtyCount++;
                                }
                            }
                        }
                    }
                } catch { }
                double avgVal = qtyCount > 0 ? qtySum / qtyCount : 0;
                formula = formula.Replace(m.Value, avgVal.ToString());
            }

            // 4. PRICE()
            Regex priceRegex = new Regex(@"PRICE\((?<cat>[^\)]+)\)");
            var priceMatches = priceRegex.Matches(formula);
            foreach (Match m in priceMatches) {
                string cat = m.Groups["cat"].Value.Trim();
                double price = GetPriceForDate(cat, dtE);
                formula = formula.Replace(m.Value, price.ToString());
            }

            // 5. 計算
            double finalVal = 0;
            try {
                DataTable dtMath = new DataTable();
                object computeResult = dtMath.Compute(formula, null);
                if (computeResult != DBNull.Value) finalVal = Convert.ToDouble(computeResult);
            } catch { finalVal = 0; }

            return finalVal;
        }

        private double GetPriceForDate(string category, DateTime date)
        {
            var matchedPrices = _prices.Where(p => p.Category == category && date >= p.StartDate && date <= p.EndDate).ToList();
            if (matchedPrices.Count > 0) return matchedPrices.First().UnitPrice;
            var fallback = _prices.Where(p => p.Category == category).OrderByDescending(p => p.EndDate).FirstOrDefault();
            if (fallback != null) return fallback.UnitPrice;
            return 0; 
        }

        // =======================================================
        // 🟢 歷史數據單價精算引擎
        // =======================================================
        private double CalculateHistoricalPrice(string formula, DateTime dtS, DateTime dtE)
        {
            string sStr = dtS.ToString("yyyy-MM-dd");
            string eStr = dtE.ToString("yyyy-MM-dd");
            string sYm = dtS.ToString("yyyy-MM");
            string eYm = dtE.ToString("yyyy-MM");

            // SUM()
            Regex sumRegex = new Regex(@"SUM\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
            var sumMatches = sumRegex.Matches(formula);
            foreach (Match m in sumMatches) {
                string db = m.Groups["db"].Value; string tb = m.Groups["tb"].Value; string col = m.Groups["col"].Value;
                double qtySum = 0;
                try {
                    var allCols = DataManager.GetColumnNames(db, tb);
                    string dateCol = allCols.Contains("日期") ? "日期" : (allCols.Contains("年月") ? "年月" : "");
                    if (!string.IsNullOrEmpty(dateCol)) {
                        string qS = dateCol == "年月" ? sYm : sStr;
                        string qE = dateCol == "年月" ? eYm : eStr;
                        DataTable dt = DataManager.GetTableData(db, tb, dateCol, qS, qE);
                        if (dt != null && dt.Columns.Contains(col)) {
                            foreach (DataRow r in dt.Rows) {
                                if (double.TryParse(r[col]?.ToString().Replace(",",""), out double qty)) qtySum += qty;
                            }
                        }
                    }
                } catch { }
                formula = formula.Replace(m.Value, qtySum.ToString());
            }

            // AVG()
            Regex avgRegex = new Regex(@"AVG\(\[(?<db>[^\]]+)\]\.\[(?<tb>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
            var avgMatches = avgRegex.Matches(formula);
            foreach (Match m in avgMatches) {
                string db = m.Groups["db"].Value; string tb = m.Groups["tb"].Value; string col = m.Groups["col"].Value;
                double qtySum = 0; int qtyCount = 0;
                try {
                    var allCols = DataManager.GetColumnNames(db, tb);
                    string dateCol = allCols.Contains("日期") ? "日期" : (allCols.Contains("年月") ? "年月" : "");
                    if (!string.IsNullOrEmpty(dateCol)) {
                        string qS = dateCol == "年月" ? sYm : sStr;
                        string qE = dateCol == "年月" ? eYm : eStr;
                        DataTable dt = DataManager.GetTableData(db, tb, dateCol, qS, qE);
                        if (dt != null && dt.Columns.Contains(col)) {
                            foreach (DataRow r in dt.Rows) {
                                if (double.TryParse(r[col]?.ToString().Replace(",",""), out double qty)) {
                                    qtySum += qty; qtyCount++;
                                }
                            }
                        }
                    }
                } catch { }
                double avgVal = qtyCount > 0 ? qtySum / qtyCount : 0;
                formula = formula.Replace(m.Value, avgVal.ToString());
            }

            double finalVal = 0;
            try {
                DataTable dtMath = new DataTable();
                object computeResult = dtMath.Compute(formula, null);
                if (computeResult != DBNull.Value) finalVal = Convert.ToDouble(computeResult);
            } catch { finalVal = 0; }

            return Math.Round(finalVal, 4); // 回傳到小數點下4位
        }

        // ==========================================
        // 浮動單價/費率管理視窗
        // ==========================================
        private void OpenPriceManager()
        {
            using (Form f = new Form { Text = "💰 費率與碳排係數管理中心", Size = new Size(750, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(10) };
                
                Label lblTop = new Label { 
                    Text = "在此設定各計價項目(如自來水、電費)或「碳排係數」於特定區間的單價。\n若為固定費率，可將結束日期設為 2099 年。日期格式請輸入: yyyy-MM-dd", 
                    Font = new Font("Microsoft JhengHei UI", 11F), 
                    AutoSize = true, 
                    Location = new Point(10, 10) 
                };

                Button btnExp = new Button { Text = "📤 匯出設定", Location = new Point(15, 55), Size = new Size(130, 35), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnImp = new Button { Text = "📥 匯入設定", Location = new Point(155, 55), Size = new Size(130, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                // 🟢 新增的歷史精算按鈕
                Button btnCalc = new Button { Text = "🧮 從歷史數據精算單價", Location = new Point(295, 55), Size = new Size(210, 35), BackColor = Color.Indigo, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };

                btnExp.Click += (s, e) => ExportPricesToExcel();
                btnImp.Click += (s, e) => { ImportPricesFromExcel(); f.DialogResult = DialogResult.OK; };

                pnlTop.Controls.Add(lblTop);
                pnlTop.Controls.Add(btnExp);
                pnlTop.Controls.Add(btnImp);
                pnlTop.Controls.Add(btnCalc);

                DataGridView dgv = new DataGridView { 
                    Dock = DockStyle.Fill, BackgroundColor = Color.WhiteSmoke, AllowUserToAddRows = true, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Microsoft JhengHei UI", 11F),
                    EditMode = DataGridViewEditMode.EditOnEnter 
                };
                
                dgv.Columns.Add("Id", "Id"); dgv.Columns["Id"].Visible = false;
                dgv.Columns.Add("Category", "計價/係數類別 (例: 電費)");
                dgv.Columns.Add("StartDate", "生效起日 (yyyy-MM-dd)");
                dgv.Columns.Add("EndDate", "生效迄日 (yyyy-MM-dd)");
                dgv.Columns.Add("UnitPrice", "單價/係數數值");

                // 綁定動態精算工具的事件
                btnCalc.Click += (s, e) => OpenDynamicPriceCalculator(dgv);

                foreach(var p in _prices) {
                    dgv.Rows.Add(p.Id, p.Category, p.StartDate.ToString("yyyy-MM-dd"), p.EndDate.ToString("yyyy-MM-dd"), p.UnitPrice);
                }

                Button btnSave = new Button { Text = "💾 儲存所有費率", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                
                btnSave.Click += (s, e) => {
                    dgv.EndEdit();
                    
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                new SQLiteCommand($"DELETE FROM {PriceTable}", conn, trans).ExecuteNonQuery();
                                
                                foreach(DataGridViewRow r in dgv.Rows) {
                                    if (r.IsNewRow) continue;
                                    string cat = r.Cells["Category"].Value?.ToString();
                                    string sd = r.Cells["StartDate"].Value?.ToString();
                                    string ed = r.Cells["EndDate"].Value?.ToString();
                                    string pr = r.Cells["UnitPrice"].Value?.ToString();

                                    if (!string.IsNullOrWhiteSpace(cat) && DateTime.TryParse(sd, out _) && DateTime.TryParse(ed, out _) && double.TryParse(pr, out double prVal)) {
                                        var cmd = new SQLiteCommand($"INSERT INTO {PriceTable} (Category, StartDate, EndDate, UnitPrice) VALUES (@c, @s, @e, @u)", conn, trans);
                                        cmd.Parameters.AddWithValue("@c", cat);
                                        cmd.Parameters.AddWithValue("@s", sd);
                                        cmd.Parameters.AddWithValue("@e", ed);
                                        cmd.Parameters.AddWithValue("@u", prVal);
                                        cmd.ExecuteNonQuery();
                                    } else {
                                        MessageBox.Show("部分資料格式錯誤，請確保日期格式為 yyyy-MM-dd，數值為數字。儲存已中止。"); 
                                        trans.Rollback();
                                        return;
                                    }
                                }
                                trans.Commit();
                            }
                        }
                        
                        LoadCache(); 
                        MessageBox.Show("費率與碳排係數儲存成功！", "成功");
                        f.DialogResult = DialogResult.OK;
                    } catch (Exception ex) {
                        MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤");
                    }
                };

                f.Controls.Add(dgv);
                f.Controls.Add(btnSave);
                f.Controls.Add(pnlTop);
                f.ShowDialog();
            }
        }

        // =======================================================
        // 🟢 動態單價試算子視窗 (加寬版面 + 範本儲存功能)
        // =======================================================
        private void OpenDynamicPriceCalculator(DataGridView dgvTarget)
        {
            using (Form f = new Form { Text = "🧮 歷史數據單價精算工具", Size = new Size(1000, 680), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, BackColor = Color.WhiteSmoke })
            {
                Label lblTop = new Label { 
                    Text = "在此設定公式，系統將會從資料庫撈取特定區間的數據進行計算，\n算出您需要的單價（例如：總繳費金額 / 總排放量 = 年度平均排放費用）。", 
                    Font = new Font("Microsoft JhengHei UI", 11F), 
                    ForeColor = Color.DimGray,
                    AutoSize = true, 
                    Location = new Point(20, 15) 
                };

                // 🟢 預設公式範本區塊
                Panel pnlTemplate = new Panel { Width = 940, Height = 45, Location = new Point(20, 65) };
                pnlTemplate.Controls.Add(new Label { Text = "歷史公式範本：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                ComboBox cboTemplates = new ComboBox { Width = 300, Location = new Point(130, 7), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button btnLoadTemplate = new Button { Text = "📥 載入", Width = 90, Height = 32, Location = new Point(440, 6), BackColor = Color.SteelBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                btnLoadTemplate.FlatAppearance.BorderSize = 0;
                Button btnSaveTemplate = new Button { Text = "💾 儲存目前公式", Width = 160, Height = 32, Location = new Point(540, 6), BackColor = Color.ForestGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                btnSaveTemplate.FlatAppearance.BorderSize = 0;
                Button btnDelTemplate = new Button { Text = "❌ 刪除範本", Width = 120, Height = 32, Location = new Point(710, 6), BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                btnDelTemplate.FlatAppearance.BorderSize = 0;
                
                pnlTemplate.Controls.AddRange(new Control[] { cboTemplates, btnLoadTemplate, btnSaveTemplate, btnDelTemplate });

                // 填寫基本資料
                Panel pnlBase = new Panel { Width = 940, Height = 45, Location = new Point(20, 120) };
                pnlBase.Controls.Add(new Label { Text = "計價類別名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtCat = new TextBox { Width = 150, Location = new Point(130, 7), Font = new Font("Microsoft JhengHei UI", 12F) };
                pnlBase.Controls.Add(txtCat);

                pnlBase.Controls.Add(new Label { Text = "萃取資料區間：", AutoSize = true, Location = new Point(310, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                DateTimePicker dtpS = new DateTimePicker { Width = 140, Location = new Point(440, 7), Format = DateTimePickerFormat.Short, Font = new Font("Microsoft JhengHei UI", 12F), Value = new DateTime(DateTime.Today.Year - 1, 1, 1) };
                pnlBase.Controls.Add(dtpS);
                pnlBase.Controls.Add(new Label { Text = "~", AutoSize = true, Location = new Point(585, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                DateTimePicker dtpE = new DateTimePicker { Width = 140, Location = new Point(610, 7), Format = DateTimePickerFormat.Short, Font = new Font("Microsoft JhengHei UI", 12F), Value = new DateTime(DateTime.Today.Year - 1, 12, 31) };
                pnlBase.Controls.Add(dtpE);

                // 公式變數生成器 (版面拉寬與調整間距)
                GroupBox boxBuilder = new GroupBox { Text = "來源變數生成器", Width=940, Height = 100, Location = new Point(20, 175), Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding=new Padding(10) };
                
                ComboBox cbAction = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(15, 45) };
                cbAction.Items.AddRange(new string[] { "SUM", "AVG" });
                cbAction.SelectedIndex = 0;

                ComboBox cbDb = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(145, 45) };
                ComboBox cbTb = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(305, 45) };
                ComboBox cbCol = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Location = new Point(545, 45) };
                
                Button btnInsert = new Button { Text = "插入變數 ⬇️", Width = 130, Height = 36, BackColor = Color.SteelBlue, ForeColor=Color.White, Cursor=Cursors.Hand, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), FlatStyle=FlatStyle.Flat, Location = new Point(795, 43) };
                btnInsert.FlatAppearance.BorderSize = 0;

                boxBuilder.Controls.Add(cbAction);
                boxBuilder.Controls.Add(new Label { Text = "(", AutoSize = true, Font = new Font("Consolas", 14F), Location = new Point(125, 45) });
                boxBuilder.Controls.Add(cbDb);
                boxBuilder.Controls.Add(new Label { Text = ".", AutoSize = true, Font = new Font("Consolas", 14F), Location = new Point(290, 45) });
                boxBuilder.Controls.Add(cbTb);
                boxBuilder.Controls.Add(new Label { Text = ".", AutoSize = true, Font = new Font("Consolas", 14F), Location = new Point(530, 45) });
                boxBuilder.Controls.Add(cbCol);
                boxBuilder.Controls.Add(new Label { Text = ")", AutoSize = true, Font = new Font("Consolas", 14F), Location = new Point(770, 45) });
                boxBuilder.Controls.Add(btnInsert);

                // 綁定中英文對照
                cbDb.Items.Add(new ItemMap { EnName="", ChName="" });
                foreach(var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName=kvp.Key, ChName=kvp.Value.ChDbName });

                cbDb.SelectedIndexChanged += (s, e) => {
                    cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName="", ChName="" });
                    var db = cbDb.SelectedItem as ItemMap;
                    if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                        foreach(var tb in _dbMap[db.EnName].Tables) cbTb.Items.Add(new ItemMap{ EnName=tb.Key, ChName=tb.Value });
                    }
                };

                cbTb.SelectedIndexChanged += (s, e) => {
                    cbCol.Items.Clear();
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                        var cols = DataManager.GetColumnNames(db.EnName, tb.EnName);
                        foreach(var c in cols.Where(x => x != "Id" && !x.Contains("日期") && !x.Contains("年月"))) cbCol.Items.Add(c);
                    }
                };

                // 計算符號與公式輸入
                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Width=940, Height = 45, Location = new Point(20, 285), WrapContents = false };
                string[] keys = { "+", "-", "*", "/", "(", ")" };
                RichTextBox rtbFormula = new RichTextBox { Width=940, Height=150, Font = new Font("Consolas", 15F), BackColor = Color.AliceBlue, Location = new Point(20, 335) };
                
                foreach (var k in keys) {
                    Button b = new Button { Text = k, Width = 50, Height = 35, Font=new Font("Consolas", 14F, FontStyle.Bold), Cursor=Cursors.Hand };
                    pnlKeys.Controls.Add(b);
                    b.Click += (s, e) => rtbFormula.AppendText(" " + b.Text + " ");
                }

                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇庫、表、欄位！"); return; }
                    rtbFormula.AppendText($"{cbAction.Text}([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}])");
                };

                // 載入與儲存範本邏輯
                Action loadTemplates = () => {
                    cboTemplates.Items.Clear();
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand($"SELECT TemplateName FROM {PriceCalcTemplateTable}", conn))
                            using (var r = cmd.ExecuteReader()) {
                                while (r.Read()) cboTemplates.Items.Add(r["TemplateName"].ToString());
                            }
                        }
                    } catch { }
                };
                loadTemplates();

                btnLoadTemplate.Click += (s, e) => {
                    if (cboTemplates.SelectedItem == null) return;
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand($"SELECT Formula FROM {PriceCalcTemplateTable} WHERE TemplateName=@t", conn)) {
                                cmd.Parameters.AddWithValue("@t", cboTemplates.SelectedItem.ToString());
                                var res = cmd.ExecuteScalar();
                                if (res != null) rtbFormula.Text = res.ToString();
                            }
                        }
                    } catch { }
                };

                btnSaveTemplate.Click += (s, e) => {
                    if (string.IsNullOrWhiteSpace(rtbFormula.Text)) { MessageBox.Show("公式不能為空！"); return; }
                    string tName = ShowInputBox("請輸入此公式範本的名稱：", "儲存範本", "新公式範本");
                    if (string.IsNullOrWhiteSpace(tName)) return;

                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            var checkCmd = new SQLiteCommand($"SELECT COUNT(1) FROM {PriceCalcTemplateTable} WHERE TemplateName=@t", conn);
                            checkCmd.Parameters.AddWithValue("@t", tName);
                            long cnt = (long)checkCmd.ExecuteScalar();

                            if (cnt > 0) {
                                var updateCmd = new SQLiteCommand($"UPDATE {PriceCalcTemplateTable} SET Formula=@f WHERE TemplateName=@t", conn);
                                updateCmd.Parameters.AddWithValue("@f", rtbFormula.Text);
                                updateCmd.Parameters.AddWithValue("@t", tName);
                                updateCmd.ExecuteNonQuery();
                            } else {
                                var insertCmd = new SQLiteCommand($"INSERT INTO {PriceCalcTemplateTable} (TemplateName, Formula) VALUES (@t, @f)", conn);
                                insertCmd.Parameters.AddWithValue("@t", tName);
                                insertCmd.Parameters.AddWithValue("@f", rtbFormula.Text);
                                insertCmd.ExecuteNonQuery();
                            }
                        }
                        MessageBox.Show("公式範本儲存成功！");
                        loadTemplates();
                        cboTemplates.SelectedItem = tName;
                    } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
                };

                btnDelTemplate.Click += (s, e) => {
                    if (cboTemplates.SelectedItem == null) return;
                    string tName = cboTemplates.SelectedItem.ToString();
                    if (MessageBox.Show($"確定要刪除範本【{tName}】？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        try {
                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                var cmd = new SQLiteCommand($"DELETE FROM {PriceCalcTemplateTable} WHERE TemplateName=@t", conn);
                                cmd.Parameters.AddWithValue("@t", tName);
                                cmd.ExecuteNonQuery();
                            }
                            loadTemplates();
                            rtbFormula.Clear();
                        } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message); }
                    }
                };

                // 底部測試與確認按鈕
                double lastTestedValue = 0;
                bool isTested = false;

                Button btnTest = new Button { Text = "⚙️ 測試計算結果", Width = 180, Height = 45, Location = new Point(20, 500), BackColor = Color.Teal, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                Label lblTestResult = new Label { Text = "試算結果：未執行", Font = new Font("Consolas", 14F, FontStyle.Bold), ForeColor = Color.Crimson, AutoSize = true, Location = new Point(210, 510) };
                
                btnTest.Click += (s, e) => {
                    if (string.IsNullOrWhiteSpace(rtbFormula.Text)) { MessageBox.Show("請先輸入公式！"); return; }
                    try {
                        lastTestedValue = CalculateHistoricalPrice(rtbFormula.Text, dtpS.Value, dtpE.Value);
                        lblTestResult.Text = $"試算結果： {lastTestedValue:N4}";
                        isTested = true;
                    } catch (Exception ex) {
                        lblTestResult.Text = "試算失敗！(可能有分母為0)";
                        MessageBox.Show("計算發生錯誤：" + ex.Message);
                    }
                };

                Button btnConfirm = new Button { Text = "✔️ 確認並帶入費率表", Width = 200, Height = 55, Location = new Point(760, 495), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnConfirm.Click += (s, e) => {
                    if (string.IsNullOrWhiteSpace(txtCat.Text)) { MessageBox.Show("請填寫計價類別名稱！"); return; }
                    if (!isTested) { MessageBox.Show("請先點擊「測試計算結果」確認數值無誤再帶入！"); return; }
                    
                    dgvTarget.Rows.Add("", txtCat.Text.Trim(), dtpS.Value.ToString("yyyy-MM-dd"), dtpE.Value.ToString("yyyy-MM-dd"), Math.Round(lastTestedValue, 4));
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(lblTop);
                f.Controls.Add(pnlTemplate);
                f.Controls.Add(pnlBase);
                f.Controls.Add(boxBuilder);
                f.Controls.Add(pnlKeys);
                f.Controls.Add(rtbFormula);
                f.Controls.Add(btnTest);
                f.Controls.Add(lblTestResult);
                f.Controls.Add(btnConfirm);

                f.ShowDialog(Form.ActiveForm);
            }
        }

        private string ShowInputBox(string prompt, string title, string defaultValue)
        {
            using (Form form = new Form { Width = 400, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = title, StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White })
            {
                Label label = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox textBox = new TextBox() { Left = 20, Top = 60, Width = 340, Text = defaultValue, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button confirmation = new Button() { Text = "確認", Left = 160, Width = 90, Height = 35, Top = 100, DialogResult = DialogResult.OK, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button cancel = new Button() { Text = "取消", Left = 270, Width = 90, Height = 35, Top = 100, DialogResult = DialogResult.Cancel, Font = new Font("Microsoft JhengHei UI", 11F) };

                form.Controls.Add(label); form.Controls.Add(textBox); form.Controls.Add(confirmation); form.Controls.Add(cancel);
                form.AcceptButton = confirmation;

                return form.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }

        private void ExportPricesToExcel()
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "費率與碳排係數設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                        {
                            conn.Open();
                            using (var cmd = new SQLiteCommand($"SELECT Category AS [類別], StartDate AS [生效起日], EndDate AS [生效迄日], UnitPrice AS [數值] FROM {PriceTable}", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("費率設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("費率設定匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportPricesFromExcel()
        {
            string authPrompt = "匯入費率設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的費率設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand($"DELETE FROM {PriceTable}", conn, trans).ExecuteNonQuery();

                                    for (int r = 2; r <= ws.Dimension.Rows; r++) 
                                    {
                                        string cat = ws.Cells[r, 1].Text.Trim();
                                        string sd = ws.Cells[r, 2].Text.Trim();
                                        string ed = ws.Cells[r, 3].Text.Trim();
                                        string pr = ws.Cells[r, 4].Text.Trim();

                                        if (string.IsNullOrEmpty(cat) || !DateTime.TryParse(sd, out _) || !DateTime.TryParse(ed, out _) || !double.TryParse(pr, out double prVal)) 
                                            continue;

                                        var cmd = new SQLiteCommand($"INSERT INTO {PriceTable} (Category, StartDate, EndDate, UnitPrice) VALUES (@c, @s, @e, @u)", conn, trans);
                                        cmd.Parameters.AddWithValue("@c", cat);
                                        cmd.Parameters.AddWithValue("@s", sd);
                                        cmd.Parameters.AddWithValue("@e", ed);
                                        cmd.Parameters.AddWithValue("@u", prVal);
                                        cmd.ExecuteNonQuery();
                                    }
                                    trans.Commit();
                                }
                            }
                            LoadCache();
                        }
                        
                        MessageBox.Show("費率設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ==========================================
        // 🟢 統一管理的公式與統計設定介面
        // ==========================================
        private void OpenUnifiedConfigManager()
        {
            using (Form f = new Form { Text = "⚙️ 統一看板公式與統計設定總管", Size = new Size(1300, 750), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340F)); 
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                // ==== 左側：清單與匯出匯入 ====
                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                
                // 🟢 加入篩選區塊下拉選單
                Panel pnlFilter = new Panel { Dock = DockStyle.Top, Height = 75, Padding = new Padding(0, 0, 0, 10) };
                Label lblFilter = new Label { Text = "1. 選擇看版區塊：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Location = new Point(0, 5) };
                ComboBox cboSection = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Location = new Point(0, 30), Width = 310 };
                cboSection.Items.AddRange(new string[] { "廢水處理", "淨水處理", "回收水", "雨水回收", "污泥減量" });
                cboSection.SelectedIndex = 0;
                pnlFilter.Controls.Add(lblFilter);
                pnlFilter.Controls.Add(cboSection);
                
                Label l1 = new Label { Text = "2. 已建立的項目 (點擊編輯)", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 35 };
                ListBox lbItems = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 45, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Margin = new Padding(0, 10, 0, 0) };
                
                Panel pnlIo = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(0, 15, 0, 5) };
                Button btnExpConf = new Button { Text = "📤 匯出全部", Width = 145, Dock = DockStyle.Left, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnImpConf = new Button { Text = "📥 匯入設定", Width = 145, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                btnExpConf.Click += (s, e) => ExportFormulasToExcel();
                btnImpConf.Click += (s, e) => { ImportFormulasFromExcel(); f.DialogResult = DialogResult.OK; };

                pnlIo.Controls.Add(btnExpConf);
                pnlIo.Controls.Add(btnImpConf);

                pnlLeft.Controls.Add(lbItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(pnlFilter);
                pnlLeft.Controls.Add(pnlIo);
                pnlLeft.Controls.Add(btnDel);
                btnDel.BringToFront(); 

                // ==== 右側：編輯區 ====
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                Label l2 = new Label { Text = "編輯 / 新增統計公式", Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Top, Height = 45 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                Panel pName = new Panel { Width = 850, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 180, Location = new Point(115, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);

                pName.Controls.Add(new Label { Text = "產出格式：", AutoSize = true, Location = new Point(320, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                ComboBox cboFormat = new ComboBox { Width = 135, Location = new Point(435, 7), Font = new Font("Microsoft JhengHei UI", 12F), DropDownStyle=ComboBoxStyle.DropDown };
                cboFormat.Items.AddRange(new string[] { "金額", "數量", "碳排(kgCO2e)", "百分比" });
                cboFormat.SelectedIndex = 0;
                pName.Controls.Add(cboFormat);
                
                Panel pName2 = new Panel { Width = 850, Height = 50 };
                pName2.Controls.Add(new Label { Text = "成本屬性：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                ComboBox cboCostType = new ComboBox { Width = 180, Location = new Point(115, 7), Font = new Font("Microsoft JhengHei UI", 12F), DropDownStyle=ComboBoxStyle.DropDownList };
                cboCostType.Items.AddRange(new string[] { "無", "支出", "效益" });
                cboCostType.SelectedIndex = 0;
                pName2.Controls.Add(cboCostType);

                pName2.Controls.Add(new Label { Text = "自訂單位：", AutoSize = true, Location = new Point(320, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                TextBox txtUnit = new TextBox { Width = 135, Location = new Point(435, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName2.Controls.Add(txtUnit);
                
                flpEditor.Controls.Add(pName);
                flpEditor.Controls.Add(pName2);

                // 產生公式區塊
                GroupBox boxBuilder = new GroupBox { Text = "公式變數生成器 (防呆選擇)", Width=850, Height = 145, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding=new Padding(10) };
                Panel pnlBuilder = new Panel { Dock = DockStyle.Fill };
                
                ComboBox cbDb = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbTb = new ComboBox { Width = 210, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbCol = new ComboBox { Width = 210, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                
                ComboBox cbAction = new ComboBox { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                cbAction.Items.AddRange(new string[] { "純數量加總 (SUM)", "平均值 (AVG)" });
                cbAction.SelectedIndex = 0;

                Button btnInsert = new Button { Text = "插入變數 ⬇️", Width = 130, Height = 36, BackColor = Color.SteelBlue, ForeColor=Color.White, Cursor=Cursors.Hand, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), FlatStyle=FlatStyle.Flat };
                btnInsert.FlatAppearance.BorderSize = 0;

                // 排版：第一排 (庫、表、欄)
                pnlBuilder.Controls.Add(new Label { Text = "庫:", Location = new Point(10, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                cbDb.Location = new Point(50, 17);
                pnlBuilder.Controls.Add(cbDb);

                pnlBuilder.Controls.Add(new Label { Text = "表:", Location = new Point(220, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                cbTb.Location = new Point(260, 17);
                pnlBuilder.Controls.Add(cbTb);

                pnlBuilder.Controls.Add(new Label { Text = "欄:", Location = new Point(490, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                cbCol.Location = new Point(530, 17);
                pnlBuilder.Controls.Add(cbCol);

                // 排版：第二排 (動作、插入按鈕)
                pnlBuilder.Controls.Add(new Label { Text = "動作:", Location = new Point(10, 68), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                cbAction.Location = new Point(65, 65);
                pnlBuilder.Controls.Add(cbAction);

                btnInsert.Location = new Point(275, 63); 
                pnlBuilder.Controls.Add(btnInsert);

                // 綁定中英文對照
                cbDb.Items.Add(new ItemMap { EnName="", ChName="" });
                foreach(var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName=kvp.Key, ChName=kvp.Value.ChDbName });

                cbDb.SelectedIndexChanged += (s, e) => {
                    cbTb.Items.Clear(); cbTb.Items.Add(new ItemMap { EnName="", ChName="" });
                    var db = cbDb.SelectedItem as ItemMap;
                    if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                        foreach(var tb in _dbMap[db.EnName].Tables) cbTb.Items.Add(new ItemMap{ EnName=tb.Key, ChName=tb.Value });
                    }
                };

                cbTb.SelectedIndexChanged += (s, e) => {
                    cbCol.Items.Clear();
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                        var cols = DataManager.GetColumnNames(db.EnName, tb.EnName);
                        foreach(var c in cols.Where(x => x != "Id" && !x.Contains("日期") && !x.Contains("年月"))) cbCol.Items.Add(c);
                    }
                };

                boxBuilder.Controls.Add(pnlBuilder);
                flpEditor.Controls.Add(boxBuilder);

                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Width=850, Height = 50, Padding = new Padding(0, 10, 0, 5), WrapContents = false };
                string[] keys = { "+", "-", "*", "/", "(", ")" };
                
                RichTextBox rtbFormula = new RichTextBox { Width=850, Height=150, Font = new Font("Consolas", 13F), BackColor = Color.AliceBlue, Margin = new Padding(0, 5, 0, 0) };
                
                foreach (var k in keys) {
                    Button b = new Button { Text = k, Width = 55, Height = 35, Font=new Font("Consolas", 14F, FontStyle.Bold), Cursor=Cursors.Hand };
                    pnlKeys.Controls.Add(b);
                    b.Click += (s, e) => rtbFormula.AppendText(" " + b.Text + " ");
                }

                pnlKeys.Controls.Add(new Panel { Width = 15, Height = 10 });

                Label lblPrice = new Label { Text = "綁定費率：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0, 5, 0, 0) };
                ComboBox cbPrice = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Margin = new Padding(0, 3, 10, 0) };
                cbPrice.Items.AddRange(_prices.Select(p => p.Category).Distinct().ToArray());
                if (cbPrice.Items.Count > 0) cbPrice.SelectedIndex = 0;

                Button btnInsertPrice = new Button { Text = "插入費率", Width = 100, Height = 35, BackColor = Color.DarkOrange, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 0, 0) };
                btnInsertPrice.FlatAppearance.BorderSize = 0;
                
                pnlKeys.Controls.Add(lblPrice);
                pnlKeys.Controls.Add(cbPrice);
                pnlKeys.Controls.Add(btnInsertPrice);

                flpEditor.Controls.Add(pnlKeys);

                Label lblF = new Label { Text = "計算公式 (可混合純數字與變數)：", Height = 30, Font=new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 10, 0, 0) };

                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇庫、表、欄位！"); return; }
                    
                    if (cbAction.SelectedIndex == 1) {
                        rtbFormula.AppendText($"AVG([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}])");
                    } else {
                        rtbFormula.AppendText($"SUM([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}])");
                    }
                };

                btnInsertPrice.Click += (s, e) => {
                    if (cbPrice.SelectedItem != null) {
                        rtbFormula.AppendText($" PRICE({cbPrice.SelectedItem}) ");
                    } else {
                        MessageBox.Show("請先選擇要綁定的費率類別！");
                    }
                };

                flpEditor.Controls.Add(lblF);
                flpEditor.Controls.Add(rtbFormula);

                Button btnSaveRow = new Button { Text = "💾 儲存並加入此區塊清單", Width = 850, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 25, 0, 0), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSaveRow.FlatAppearance.BorderSize = 0;

                pnlRight.Controls.Add(flpEditor);
                pnlRight.Controls.Add(l2);
                pnlRight.Controls.Add(btnSaveRow);
                btnSaveRow.Dock = DockStyle.Bottom;

                tlp.Controls.Add(pnlLeft, 0, 0);
                tlp.Controls.Add(pnlRight, 1, 0);
                f.Controls.Add(tlp);

                // 資料綁定
                Action refreshList = () => {
                    lbItems.Items.Clear();
                    string sec = cboSection.Text;
                    foreach (var cfg in _configs.Where(x => x.Section == sec)) {
                        lbItems.Items.Add(cfg.DisplayName);
                    }
                };

                cboSection.SelectedIndexChanged += (s, e) => {
                    refreshList();
                    txtName.Clear(); rtbFormula.Clear(); txtUnit.Clear(); cboFormat.SelectedIndex = 0; cboCostType.SelectedIndex = 0;
                };

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (lbItems.SelectedIndex < 0) return;
                    string sec = cboSection.Text;
                    var cfg = _configs.First(x => x.Section == sec && x.DisplayName == lbItems.SelectedItem.ToString());
                    txtName.Text = cfg.DisplayName;
                    cboFormat.Text = cfg.OutputType;
                    cboCostType.Text = cfg.CostType; 
                    txtUnit.Text = cfg.Unit;
                    rtbFormula.Text = cfg.Formula;
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        string sec = cboSection.Text;
                        _configs.RemoveAll(x => x.Section == sec && x.DisplayName == lbItems.SelectedItem.ToString());
                        SaveConfigsToDb(); refreshList(); txtName.Clear(); rtbFormula.Clear(); txtUnit.Clear();
                    }
                };

                btnSaveRow.Click += (ss, ee) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(rtbFormula.Text)) { MessageBox.Show("請輸入顯示名稱與公式！"); return; }
                    
                    string sec = cboSection.Text;
                    _configs.RemoveAll(x => x.Section == sec && x.DisplayName == txtName.Text);
                    _configs.Add(new CostFormulaItem { 
                        Section = sec, 
                        DisplayName = txtName.Text.Trim(), 
                        OutputType = cboFormat.Text, 
                        CostType = cboCostType.Text, 
                        Unit = txtUnit.Text.Trim(), 
                        Formula = rtbFormula.Text.Trim() 
                    });
                    
                    SaveConfigsToDb(); refreshList();
                    MessageBox.Show("儲存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                f.ShowDialog();
            }
        }

        private void SaveConfigsToDb()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        // 清除資料表所有資料
                        new SQLiteCommand($"DELETE FROM {ConfigTable}", conn, trans).ExecuteNonQuery();
                        
                        // 重新寫入最新的所有資料
                        foreach(var c in _configs) {
                            var cmd = new SQLiteCommand($"INSERT INTO {ConfigTable} (Section, DisplayName, OutputType, CostType, Unit, Formula) VALUES (@s, @d, @o, @ct, @u, @f)", conn, trans);
                            cmd.Parameters.AddWithValue("@s", c.Section);
                            cmd.Parameters.AddWithValue("@d", c.DisplayName);
                            cmd.Parameters.AddWithValue("@o", c.OutputType);
                            cmd.Parameters.AddWithValue("@ct", c.CostType);
                            cmd.Parameters.AddWithValue("@u", c.Unit);
                            cmd.Parameters.AddWithValue("@f", c.Formula);
                            cmd.ExecuteNonQuery();
                        }
                        trans.Commit();
                    }
                }
                LoadCache(); // 重載確保同步
            } catch (Exception ex) { 
                MessageBox.Show($"儲存配置失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportFormulasToExcel()
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "成本統計公式設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                        {
                            conn.Open();
                            using (var cmd = new SQLiteCommand($"SELECT Section AS [看板區塊], DisplayName AS [顯示名稱], OutputType AS [產出格式], CostType AS [成本屬性(支出/效益)], Unit AS [自訂單位], Formula AS [計算公式] FROM {ConfigTable}", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("成本公式設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("公式設定匯出成功！您可以以此檔案作為備份。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportFormulasFromExcel()
        {
            string authPrompt = "匯入公式設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的公式設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand($"DELETE FROM {ConfigTable}", conn, trans).ExecuteNonQuery();

                                    for (int r = 2; r <= ws.Dimension.Rows; r++) 
                                    {
                                        string section = ws.Cells[r, 1].Text.Trim();
                                        string disp = ws.Cells[r, 2].Text.Trim();
                                        string output = ws.Cells[r, 3].Text.Trim();
                                        string costType = ws.Cells[r, 4].Text.Trim();
                                        string unit = ws.Cells[r, 5].Text.Trim();
                                        string formula = ws.Cells[r, 6].Text.Trim();

                                        if (string.IsNullOrEmpty(section) || string.IsNullOrEmpty(disp) || string.IsNullOrEmpty(formula)) continue;

                                        var cmd = new SQLiteCommand($"INSERT INTO {ConfigTable} (Section, DisplayName, OutputType, CostType, Unit, Formula) VALUES (@s, @d, @o, @ct, @u, @f)", conn, trans);
                                        cmd.Parameters.AddWithValue("@s", section);
                                        cmd.Parameters.AddWithValue("@d", disp);
                                        cmd.Parameters.AddWithValue("@o", output);
                                        cmd.Parameters.AddWithValue("@ct", costType);
                                        cmd.Parameters.AddWithValue("@u", unit);
                                        cmd.Parameters.AddWithValue("@f", formula);
                                        cmd.ExecuteNonQuery();
                                    }
                                    trans.Commit();
                                }
                            }
                            LoadCache();
                        }
                        
                        MessageBox.Show("公式設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ExecuteCalculation(); 
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ==========================================
        // 導出 PDF 功能與附錄產生器
        // ==========================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 400, Height = 460, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Top, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Top, Height = 240, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 14F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                clb.Items.Add("廢水處理費用統計", true); 
                clb.Items.Add("淨水處理費用統計", true); 
                clb.Items.Add("回收水成本與效益", true);
                clb.Items.Add("雨水回收效益分析", true); 
                clb.Items.Add("污泥減量與處置效益", true); 
                
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    if (clb.GetItemChecked(0)) selectedPanels.Add(_sec1.MainBox);
                    if (clb.GetItemChecked(1)) selectedPanels.Add(_sec2.MainBox);
                    if (clb.GetItemChecked(2)) selectedPanels.Add(_sec3.MainBox);
                    if (clb.GetItemChecked(3)) selectedPanels.Add(_sec4.MainBox); 
                    if (clb.GetItemChecked(4)) selectedPanels.Add(_sec5.MainBox); 
                }
            }
            return selectedPanels;
        }

        private void ExportToPdf()
        {
            var panelsToExport = GetSelectedExportPanels();
            if (panelsToExport.Count == 0) return;

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try 
            {
                foreach (Control ctrl in _controlsToHideForPdf) ctrl.Visible = false;
                Application.DoEvents(); 

                List<Bitmap> bitmaps = new List<Bitmap>();
                foreach (var pnl in panelsToExport) 
                {
                    Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                    pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                    bitmaps.Add(bmp);
                }

                Bitmap appendixBmp = CreateAppendixBitmap();
                bitmaps.Add(appendixBmp);

                string dateStr = $"結算區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text} ~ {_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
                
                PdfHelper.ExportDashboardToPdf(bitmaps, "水資源成本與 ESG 效益分析報表", dateStr, "水資源成本與效益分析報表");
            } 
            catch (Exception ex)
            {
                MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (Control ctrl in _controlsToHideForPdf) ctrl.Visible = true;
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private Bitmap CreateAppendixBitmap()
        {
            int width = 1100;
            int height = 200 + (_configs.Count * 45) + (_prices.Count * 35);
            if (height < 600) height = 600;

            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fSection = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
                Font fItem = new Font("Consolas", 12F);
                Font fNote = new Font("Microsoft JhengHei UI", 10F);

                float y = 30;
                float x = 40;

                g.DrawString("【附錄】ESG 數據溯源與計算公式總覽", fTitle, Brushes.DarkSlateBlue, x, y);
                y += 45;
                g.DrawString("此附錄為確保數據可追溯性 (Traceability) 而自動生成，詳列報表中使用的所有計算公式與費率係數。", fNote, Brushes.DimGray, x, y);
                y += 50;

                // 1. 公式清單
                g.DrawString("一、 統計公式設定清單", fSection, Brushes.Teal, x, y);
                y += 35;

                foreach (var cfg in _configs.OrderBy(c => c.Section))
                {
                    string outputMark = cfg.OutputType == "金額" ? $"[{cfg.CostType}]" : (cfg.OutputType == "碳排(kgCO2e)" ? "[☁️]" : "[#]");
                    string line = $"[{cfg.Section}] {cfg.DisplayName} {outputMark} = {cfg.Formula}";
                    g.DrawString(line, fItem, Brushes.Black, x + 20, y);
                    y += 30;
                }

                y += 30;

                // 2. 費率清單
                g.DrawString("二、 費率與碳排係數清單", fSection, Brushes.DarkOrange, x, y);
                y += 35;

                foreach (var p in _prices.OrderBy(p => p.Category).ThenBy(p => p.StartDate))
                {
                    string endStr = p.EndDate.Year >= 2099 ? "長期有效" : p.EndDate.ToString("yyyy/MM/dd");
                    string line = $"[{p.Category}] 適用區間: {p.StartDate:yyyy/MM/dd} ~ {endStr} => 數值: {p.UnitPrice}";
                    g.DrawString(line, fItem, Brushes.Black, x + 20, y);
                    y += 30;
                }
            }

            return bmp;
        }

        private void InitDateComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) { _cboStartYear.Items.Add(i); _cboEndYear.Items.Add(i); }
            for (int i = 1; i <= 12; i++) { _cboStartMonth.Items.Add(i.ToString("D2")); _cboEndMonth.Items.Add(i.ToString("D2")); }
            
            _cboStartYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboStartMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboEndYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            _cboEndMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            
            DateTime today = DateTime.Today;
            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, new DateTime(today.Year, today.Month, 1));
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

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            y.SelectedItem = date.Year; m.SelectedItem = date.Month.ToString("D2");
            UpdateDaysCombo(y, m, d); d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetDateFromCombo(ComboBox y, ComboBox m, ComboBox d) {
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            return new DateTime((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()), day > maxDay ? maxDay : day);
        }
    }
}
