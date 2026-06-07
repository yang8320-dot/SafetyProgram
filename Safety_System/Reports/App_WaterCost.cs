using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterCost
    {
        // 頂部日期區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        private FlowLayoutPanel _flpSection1, _flpSection2, _flpSection3;
        private Label _lblTotal1, _lblTotal2, _lblTotal3; // 區塊總計
        
        // 🟢 PDF 匯出用區塊外框
        private Panel _pnlBox1, _pnlBox2, _pnlBox3;
        
        private Button _btnSearch;

        // 資料庫與快取
        private const string SysDbName = "SystemConfig";
        private const string ConfigTable = "WaterCostFormulas"; // 全新支援公式的設定表
        private const string PriceTable = "WaterPrices";

        private List<CostFormulaItem> _configs = new List<CostFormulaItem>();
        private List<PriceItem> _prices = new List<PriceItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 模型定義
        private class CostFormulaItem {
            public int Id { get; set; }
            public string Section { get; set; }     // 廢水處理, 淨水處理, 回收水
            public string DisplayName { get; set; }
            public string OutputType { get; set; }  // 金額, 數量
            public string Formula { get; set; }     // 儲存實際計算公式
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
            TableLayoutPanel masterLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 5, Margin = new Padding(0) };
            for(int i=0; i<5; i++) masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ==========================================
            // 1. 標題與操作區
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "💰 水資源成本與效益分析看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
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

            Button btnPriceManager = new Button { Text = "💰 浮動單價/費率管理", Size = new Size(210, 42), BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnPriceManager.Click += (s, e) => { OpenPriceManager(); ExecuteCalculation(); };

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
                _btnSearch, btnPriceManager, btnPdf
            });

            // ==========================================
            // 2. 建立三大區塊
            // ==========================================
            _pnlBox1 = BuildSection("廢水處理費用統計", "廢水處理", Color.Sienna, out _flpSection1, out _lblTotal1);
            _pnlBox2 = BuildSection("淨水處理費用統計", "淨水處理", Color.MediumBlue, out _flpSection2, out _lblTotal2);
            _pnlBox3 = BuildSection("回收水成本與效益", "回收水", Color.ForestGreen, out _flpSection3, out _lblTotal3);

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);
            masterLayout.Controls.Add(_pnlBox1, 0, 2);
            masterLayout.Controls.Add(_pnlBox2, 0, 3);
            masterLayout.Controls.Add(_pnlBox3, 0, 4);

            mainScrollPanel.Controls.Add(masterLayout);

            ExecuteCalculation();

            return mainScrollPanel;
        }

        // ==========================================
        // 資料庫初始化與快取
        // ==========================================
        private void InitDatabase()
        {
            string sql1 = $"CREATE TABLE IF NOT EXISTS [{ConfigTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, Section TEXT, DisplayName TEXT, OutputType TEXT, Formula TEXT);";
            string sql2 = $"CREATE TABLE IF NOT EXISTS [{PriceTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, Category TEXT, StartDate TEXT, EndDate TEXT, UnitPrice REAL);";
            
            DataManager.InitTable(SysDbName, ConfigTable, sql1);
            DataManager.InitTable(SysDbName, PriceTable, sql2);
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
                        while (r.Read()) _configs.Add(new CostFormulaItem { 
                            Id = Convert.ToInt32(r["Id"]), Section = r["Section"].ToString(), DisplayName = r["DisplayName"].ToString(), 
                            OutputType = r["OutputType"].ToString(), Formula = r["Formula"].ToString()
                        });
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
        private Panel BuildSection(string title, string sectionCode, Color themeColor, out FlowLayoutPanel flpData, out Label lblTotal)
        {
            Panel pnlWrapper = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 25) };
            pnlWrapper.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlWrapper.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.White };
            Label lblTitle = new Label { Text = $"■ {title}", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = themeColor, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, AutoSize = true, Padding=new Padding(10,15,0,0) };
            
            lblTotal = new Label { Text = "區塊總計: $ 0", Font = new Font("Consolas", 18F, FontStyle.Bold), ForeColor = Color.Crimson, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Right, AutoSize = true, Padding=new Padding(0,10,20,0) };

            Button btnSetting = new Button { Text = "⚙️ 公式與統計設定", Size = new Size(160, 35), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, Dock = DockStyle.Right, Margin = new Padding(0,10,15,0) };
            btnSetting.Click += (s, e) => { OpenConfigManager(sectionCode); ExecuteCalculation(); };

            pnlHeader.Controls.Add(btnSetting);
            pnlHeader.Controls.Add(lblTotal);
            pnlHeader.Controls.Add(lblTitle);

            flpData = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, MinimumSize=new Size(0, 120), Padding = new Padding(15), WrapContents = true };

            pnlWrapper.Controls.Add(flpData);
            pnlWrapper.Controls.Add(pnlHeader);

            return pnlWrapper;
        }

        // ==========================================
        // 核心計算引擎 (正規表示式解析公式)
        // ==========================================
        private void ExecuteCalculation()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            RenderSectionCards("廢水處理", _flpSection1, _lblTotal1, dtS, dtE, Color.Sienna);
            RenderSectionCards("淨水處理", _flpSection2, _lblTotal2, dtS, dtE, Color.MediumBlue);
            RenderSectionCards("回收水", _flpSection3, _lblTotal3, dtS, dtE, Color.ForestGreen);

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private void RenderSectionCards(string sectionCode, FlowLayoutPanel flp, Label lblTotal, DateTime dtS, DateTime dtE, Color themeColor)
        {
            flp.Controls.Clear();
            var sectionConfigs = _configs.Where(c => c.Section == sectionCode).ToList();

            if (sectionConfigs.Count == 0) {
                flp.Controls.Add(new Label { Text = "尚未設定任何統計項目，請點擊右上角設定。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                lblTotal.Text = "";
                return;
            }

            double sectionTotalCost = 0;

            foreach (var cfg in sectionConfigs)
            {
                double computedVal = EvaluateFormula(cfg.Formula, dtS, dtE);

                Panel card = new Panel { Size = new Size(300, 110), BackColor = Color.WhiteSmoke, Margin = new Padding(10) };
                card.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, card.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                
                Label lTitle = new Label { Text = cfg.DisplayName, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.FromArgb(60,60,60), Location = new Point(15,15), AutoSize = true };
                
                Label lVal = new Label { Font = new Font("Consolas", 20F, FontStyle.Bold), Location = new Point(15, 55), AutoSize = true };

                if (cfg.OutputType == "金額") {
                    lVal.Text = $"$ {computedVal:N0}";
                    lVal.ForeColor = themeColor;
                    sectionTotalCost += computedVal;
                } else {
                    lVal.Text = $"{computedVal:N2}";
                    lVal.ForeColor = Color.DarkSlateBlue;
                }

                card.Controls.AddRange(new Control[] { lTitle, lVal });
                flp.Controls.Add(card);
            }

            lblTotal.Text = $"區塊總計金額: $ {sectionTotalCost:N0}";
        }

        private double EvaluateFormula(string formula, DateTime dtS, DateTime dtE)
        {
            string sStr = dtS.ToString("yyyy-MM-dd");
            string eStr = dtE.ToString("yyyy-MM-dd");
            string sYm = dtS.ToString("yyyy-MM");
            string eYm = dtE.ToString("yyyy-MM");

            // 1. 解析並取代 COST([DB].[TB].[Col], Category)
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

            // 2. 解析並取代 SUM([DB].[TB].[Col])
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

            // 3. 利用 DataTable 內建引擎計算數學算式 (如 1000 / (50 * 12.5))
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

        // ==========================================
        // 浮動單價/費率管理視窗
        // ==========================================
        private void OpenPriceManager()
        {
            using (Form f = new Form { Text = "💰 浮動單價/費率管理中心", Size = new Size(700, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                Label lblTop = new Label { Text = "在此設定各計價項目(如自來水、電費、藥劑)於特定區間的單價。\n若為固定費率，可將結束日期設為 2099 年。", Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(15), Dock = DockStyle.Top, Height=60 };

                DataGridView dgv = new DataGridView { 
                    Dock = DockStyle.Fill, BackgroundColor = Color.WhiteSmoke, AllowUserToAddRows = true, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Microsoft JhengHei UI", 11F) 
                };
                
                dgv.Columns.Add("Id", "Id"); dgv.Columns["Id"].Visible = false;
                dgv.Columns.Add("Category", "計價類別 (例: 電費)");
                dgv.Columns.Add("StartDate", "生效起日 (yyyy-MM-dd)");
                dgv.Columns.Add("EndDate", "生效迄日 (yyyy-MM-dd)");
                dgv.Columns.Add("UnitPrice", "單價數值");

                foreach(var p in _prices) {
                    dgv.Rows.Add(p.Id, p.Category, p.StartDate.ToString("yyyy-MM-dd"), p.EndDate.ToString("yyyy-MM-dd"), p.UnitPrice);
                }

                Button btnSave = new Button { Text = "💾 儲存所有費率", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                btnSave.Click += (s, e) => {
                    dgv.EndEdit();
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int)); dt.Columns.Add("Category", typeof(string)); dt.Columns.Add("StartDate", typeof(string)); dt.Columns.Add("EndDate", typeof(string)); dt.Columns.Add("UnitPrice", typeof(double));

                    foreach(DataGridViewRow r in dgv.Rows) {
                        if (r.IsNewRow) continue;
                        string cat = r.Cells["Category"].Value?.ToString();
                        string sd = r.Cells["StartDate"].Value?.ToString();
                        string ed = r.Cells["EndDate"].Value?.ToString();
                        string pr = r.Cells["UnitPrice"].Value?.ToString();

                        if (!string.IsNullOrWhiteSpace(cat) && DateTime.TryParse(sd, out _) && DateTime.TryParse(ed, out _) && double.TryParse(pr, out double prVal)) {
                            DataRow dr = dt.NewRow();
                            if (r.Cells["Id"].Value != null && int.TryParse(r.Cells["Id"].Value.ToString(), out int id)) dr["Id"] = id;
                            dr["Category"] = cat; dr["StartDate"] = sd; dr["EndDate"] = ed; dr["UnitPrice"] = prVal;
                            dt.Rows.Add(dr);
                        } else {
                            MessageBox.Show("部分資料格式錯誤，請確保日期格式為 yyyy-MM-dd，單價為數字。"); return;
                        }
                    }

                    DataManager.DropTable(SysDbName, PriceTable);
                    InitDatabase();
                    DataManager.BulkSaveTable(SysDbName, PriceTable, dt);
                    LoadCache();
                    MessageBox.Show("費率儲存成功！", "成功");
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(dgv);
                f.Controls.Add(btnSave);
                f.Controls.Add(lblTop);
                f.ShowDialog();
            }
        }

        // ==========================================
        // 🟢 全新公式設定介面 (支援防呆與中文顯示)
        // ==========================================
        private void OpenConfigManager(string sectionCode)
        {
            using (Form f = new Form { Text = $"⚙️ 統計項目與公式設定 ({sectionCode})", Size = new Size(1100, 680), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                // ==== 左側：清單 ====
                Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                Label l1 = new Label { Text = "已建立的統計項目", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 30 };
                ListBox lbItems = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                Button btnDel = new Button { Text = "❌ 刪除選取項目", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                pnlLeft.Controls.Add(lbItems);
                pnlLeft.Controls.Add(l1);
                pnlLeft.Controls.Add(btnDel);

                // ==== 右側：編輯區 ====
                Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
                Label l2 = new Label { Text = "編輯 / 新增統計公式", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Top, Height = 40 };

                FlowLayoutPanel flpEditor = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                Panel pName = new Panel { Width = 750, Height = 45 };
                pName.Controls.Add(new Label { Text = "顯示名稱：", AutoSize = true, Location = new Point(0, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) });
                TextBox txtName = new TextBox { Width = 250, Location = new Point(100, 7), Font = new Font("Microsoft JhengHei UI", 12F) }; 
                pName.Controls.Add(txtName);

                pName.Controls.Add(new Label { Text = "產出格式：", AutoSize = true, Location = new Point(370, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) }); 
                ComboBox cboFormat = new ComboBox { Width = 120, Location = new Point(470, 7), Font = new Font("Microsoft JhengHei UI", 12F), DropDownStyle=ComboBoxStyle.DropDownList };
                cboFormat.Items.AddRange(new string[] { "金額", "數量" });
                cboFormat.SelectedIndex = 0;
                pName.Controls.Add(cboFormat);
                
                flpEditor.Controls.Add(pName);

                // 產生公式區塊
                GroupBox boxBuilder = new GroupBox { Text = "公式變數生成器 (防呆選擇)", Width=740, Height = 140, Font=new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Padding=new Padding(10) };
                FlowLayoutPanel flpBuilder = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = true };
                
                ComboBox cbDb = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbTb = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cbCol = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                
                ComboBox cbAction = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F) };
                cbAction.Items.AddRange(new string[] { "純數量加總 (SUM)", "結合費率計算成本 (COST)" });
                cbAction.SelectedIndex = 0;

                ComboBox cbPrice = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font=new Font("Microsoft JhengHei UI", 11F), Enabled=false };
                cbPrice.Items.AddRange(_prices.Select(p => p.Category).Distinct().ToArray());

                Button btnInsert = new Button { Text = "插入變數 ⬇️", Width = 120, Height = 35, BackColor = Color.SteelBlue, ForeColor=Color.White, Cursor=Cursors.Hand };

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

                cbAction.SelectedIndexChanged += (s, e) => { cbPrice.Enabled = cbAction.SelectedIndex == 1; };

                flpBuilder.Controls.AddRange(new Control[] { new Label{Text="庫:",AutoSize=true}, cbDb, new Label{Text="表:",AutoSize=true}, cbTb, new Label{Text="欄:",AutoSize=true}, cbCol, new Label{Text="動作:",AutoSize=true}, cbAction, new Label{Text="綁定費率:",AutoSize=true}, cbPrice, btnInsert });
                boxBuilder.Controls.Add(flpBuilder);
                flpEditor.Controls.Add(boxBuilder);

                // 計算符號
                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Width=740, Height = 45, Padding = new Padding(0, 5, 0, 5) };
                string[] keys = { "+", "-", "*", "/", "(", ")" };
                foreach (var k in keys) {
                    Button b = new Button { Text = k, Width = 50, Height = 35, Font=new Font("Consolas", 14F, FontStyle.Bold) };
                    pnlKeys.Controls.Add(b);
                }
                flpEditor.Controls.Add(pnlKeys);

                RichTextBox rtbFormula = new RichTextBox { Width=740, Height=150, Font = new Font("Consolas", 13F), BackColor = Color.AliceBlue };
                Label lblF = new Label { Text = "計算公式 (可混合純數字與變數)：", Height = 25, Font=new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };

                foreach (Control c in pnlKeys.Controls) {
                    if (c is Button b) b.Click += (s, e) => rtbFormula.AppendText(" " + b.Text + " ");
                }

                btnInsert.Click += (s, e) => {
                    var db = cbDb.SelectedItem as ItemMap; var tb = cbTb.SelectedItem as ItemMap;
                    if (db == null || tb == null || cbCol.SelectedItem == null) { MessageBox.Show("請選擇庫、表、欄位！"); return; }
                    
                    if (cbAction.SelectedIndex == 1) {
                        if (cbPrice.SelectedItem == null) { MessageBox.Show("請選擇要綁定的費率類別！"); return; }
                        rtbFormula.AppendText($"COST([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}], {cbPrice.SelectedItem})");
                    } else {
                        rtbFormula.AppendText($"SUM([{db.EnName}].[{tb.EnName}].[{cbCol.SelectedItem}])");
                    }
                };

                flpEditor.Controls.Add(lblF);
                flpEditor.Controls.Add(rtbFormula);

                Button btnSaveRow = new Button { Text = "💾 儲存並加入清單", Width = 740, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 20, 0, 0), Cursor = Cursors.Hand };

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
                    foreach (var cfg in _configs.Where(x => x.Section == sectionCode)) {
                        lbItems.Items.Add(cfg.DisplayName);
                    }
                };
                refreshList();

                lbItems.SelectedIndexChanged += (ss, ee) => {
                    if (lbItems.SelectedIndex < 0) return;
                    var cfg = _configs.First(x => x.Section == sectionCode && x.DisplayName == lbItems.SelectedItem.ToString());
                    txtName.Text = cfg.DisplayName;
                    cboFormat.Text = cfg.OutputType;
                    rtbFormula.Text = cfg.Formula;
                };

                btnDel.Click += (ss, ee) => {
                    if (lbItems.SelectedIndex >= 0) {
                        _configs.RemoveAll(x => x.Section == sectionCode && x.DisplayName == lbItems.SelectedItem.ToString());
                        SaveConfigsToDb(); refreshList(); txtName.Clear(); rtbFormula.Clear();
                    }
                };

                btnSaveRow.Click += (ss, ee) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(rtbFormula.Text)) { MessageBox.Show("請輸入顯示名稱與公式！"); return; }
                    
                    _configs.RemoveAll(x => x.Section == sectionCode && x.DisplayName == txtName.Text);
                    _configs.Add(new CostFormulaItem { Section = sectionCode, DisplayName = txtName.Text.Trim(), OutputType = cboFormat.Text, Formula = rtbFormula.Text.Trim() });
                    
                    SaveConfigsToDb(); refreshList();
                    MessageBox.Show("儲存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                f.ShowDialog();
            }
        }

        private void SaveConfigsToDb()
        {
            try {
                DataTable dt = new DataTable();
                dt.Columns.Add("Section", typeof(string)); dt.Columns.Add("DisplayName", typeof(string)); 
                dt.Columns.Add("OutputType", typeof(string)); dt.Columns.Add("Formula", typeof(string));
                
                foreach(var c in _configs) {
                    DataRow r = dt.NewRow();
                    r["Section"] = c.Section; r["DisplayName"] = c.DisplayName; r["OutputType"] = c.OutputType; r["Formula"] = c.Formula;
                    dt.Rows.Add(r);
                }

                DataManager.DropTable(SysDbName, ConfigTable);
                InitDatabase();
                DataManager.BulkSaveTable(SysDbName, ConfigTable, dt);
                LoadCache();
            } catch { }
        }

        // ==========================================
        // 🟢 導出 PDF 功能與輔助方法
        // ==========================================
        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 400, Height = 350, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Top, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Top, Height = 180, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 14F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                clb.Items.Add("廢水處理費用統計", true); 
                clb.Items.Add("淨水處理費用統計", true); 
                clb.Items.Add("回收水成本與效益", true);
                
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    if (clb.GetItemChecked(0)) selectedPanels.Add(_pnlBox1);
                    if (clb.GetItemChecked(1)) selectedPanels.Add(_pnlBox2);
                    if (clb.GetItemChecked(2)) selectedPanels.Add(_pnlBox3);
                }
            }
            return selectedPanels;
        }

        private void ExportToPdf()
        {
            var panelsToExport = GetSelectedExportPanels();
            if (panelsToExport.Count == 0) return;

            List<Bitmap> bitmaps = new List<Bitmap>();
            foreach (var pnl in panelsToExport) 
            {
                // 將每個勾選的區塊截圖
                Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                bitmaps.Add(bmp);
            }

            string dateStr = $"結算區間：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text} ~ {_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}";
            
            // 呼叫 PdfHelper 的共用儀表板匯出引擎
            PdfHelper.ExportDashboardToPdf(bitmaps, "水資源成本與效益分析報表", dateStr, "水資源成本分析表");
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
