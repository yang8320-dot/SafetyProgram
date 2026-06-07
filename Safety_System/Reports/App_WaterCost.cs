using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterCost
    {
        // 頂部日期區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        private FlowLayoutPanel _flpSection1, _flpSection2, _flpSection3;
        private Button _btnSearch;

        // 資料庫與快取
        private const string SysDbName = "SystemConfig";
        private const string ConfigTable = "WaterCostConfigs";
        private const string PriceTable = "WaterPrices";

        private List<CostConfigItem> _configs = new List<CostConfigItem>();
        private List<PriceItem> _prices = new List<PriceItem>();

        // 模型定義
        private class CostConfigItem {
            public int Id { get; set; }
            public string Section { get; set; } // 廢水, 淨水, 回收水
            public string DisplayName { get; set; }
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string ColName { get; set; }
            public string PriceCategory { get; set; } // 綁定的計價項目
        }

        private class PriceItem {
            public int Id { get; set; }
            public string Category { get; set; } // 例如: 自來水費, 電費, PAC
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

            _btnSearch = new Button { Text = "🔍 計算成本", Size = new Size(130, 42), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.Click += (s, e) => ExecuteCalculation();

            Button btnPriceManager = new Button { Text = "💰 浮動單價/費率管理", Size = new Size(210, 42), BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(20, 0, 0, 0) };
            btnPriceManager.Click += (s, e) => { OpenPriceManager(); ExecuteCalculation(); };

            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(140, 42), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
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
            var sec1 = BuildSection("廢水處理費用統計", "廢水", Color.Sienna, out _flpSection1);
            var sec2 = BuildSection("淨水處理費用統計", "淨水", Color.MediumBlue, out _flpSection2);
            var sec3 = BuildSection("回收水成本與效益", "回收水", Color.ForestGreen, out _flpSection3);

            masterLayout.Controls.Add(pnlHeader, 0, 0);
            masterLayout.Controls.Add(flpControls, 0, 1);
            masterLayout.Controls.Add(sec1, 0, 2);
            masterLayout.Controls.Add(sec2, 0, 3);
            masterLayout.Controls.Add(sec3, 0, 4);

            mainScrollPanel.Controls.Add(masterLayout);

            ExecuteCalculation();

            return mainScrollPanel;
        }

        // ==========================================
        // 初始化與資料庫
        // ==========================================
        private void InitDatabase()
        {
            string sql1 = $"CREATE TABLE IF NOT EXISTS [{ConfigTable}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, Section TEXT, DisplayName TEXT, DbName TEXT, TableName TEXT, ColName TEXT, PriceCategory TEXT);";
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
                        while (r.Read()) _configs.Add(new CostConfigItem { 
                            Id = Convert.ToInt32(r["Id"]), Section = r["Section"].ToString(), DisplayName = r["DisplayName"].ToString(), 
                            DbName = r["DbName"].ToString(), TableName = r["TableName"].ToString(), ColName = r["ColName"].ToString(), PriceCategory = r["PriceCategory"].ToString() 
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
        private Panel BuildSection(string title, string sectionCode, Color themeColor, out FlowLayoutPanel flpData)
        {
            Panel pnlWrapper = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 25) };
            pnlWrapper.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlWrapper.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.White };
            Label lblTitle = new Label { Text = $"■ {title}", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = themeColor, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, AutoSize = true, Padding=new Padding(10,15,0,0) };
            
            Button btnSetting = new Button { Text = "⚙️ 統計項目設定", Size = new Size(150, 35), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, Dock = DockStyle.Right, Margin = new Padding(0,10,15,0) };
            btnSetting.Click += (s, e) => { OpenConfigManager(sectionCode); ExecuteCalculation(); };

            pnlHeader.Controls.Add(btnSetting);
            pnlHeader.Controls.Add(lblTitle);

            flpData = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, MinimumSize=new Size(0, 120), Padding = new Padding(15), WrapContents = true };

            pnlWrapper.Controls.Add(flpData);
            pnlWrapper.Controls.Add(pnlHeader);

            return pnlWrapper;
        }

        // ==========================================
        // 計算核心引擎
        // ==========================================
        private void ExecuteCalculation()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            RenderSectionCards("廢水", _flpSection1, dtS, dtE, Color.Sienna);
            RenderSectionCards("淨水", _flpSection2, dtS, dtE, Color.MediumBlue);
            RenderSectionCards("回收水", _flpSection3, dtS, dtE, Color.ForestGreen);

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private void RenderSectionCards(string sectionCode, FlowLayoutPanel flp, DateTime dtS, DateTime dtE, Color themeColor)
        {
            flp.Controls.Clear();
            var sectionConfigs = _configs.Where(c => c.Section == sectionCode).ToList();

            if (sectionConfigs.Count == 0) {
                flp.Controls.Add(new Label { Text = "尚未設定任何統計項目，請點擊右上角設定。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                return;
            }

            string sStr = dtS.ToString("yyyy-MM-dd");
            string eStr = dtE.ToString("yyyy-MM-dd");
            string sYm = dtS.ToString("yyyy-MM");
            string eYm = dtE.ToString("yyyy-MM");

            double sectionTotalCost = 0;

            foreach (var cfg in sectionConfigs)
            {
                double totalQty = 0;
                double totalCost = 0;

                try {
                    var cols = DataManager.GetColumnNames(cfg.DbName, cfg.TableName);
                    if (cols.Contains(cfg.ColName)) 
                    {
                        string dateCol = cols.Contains("日期") ? "日期" : (cols.Contains("年月") ? "年月" : "");
                        if (!string.IsNullOrEmpty(dateCol)) 
                        {
                            string qS = dateCol == "年月" ? sYm : sStr;
                            string qE = dateCol == "年月" ? eYm : eStr;

                            DataTable dt = DataManager.GetTableData(cfg.DbName, cfg.TableName, dateCol, qS, qE);
                            if (dt != null) 
                            {
                                foreach (DataRow r in dt.Rows) 
                                {
                                    if (double.TryParse(r[cfg.ColName]?.ToString().Replace(",",""), out double qty)) 
                                    {
                                        totalQty += qty;

                                        // 動態取得該日期的單價
                                        DateTime rowDate = dtS; 
                                        if (dateCol == "日期") DateTime.TryParse(r["日期"].ToString(), out rowDate);
                                        else if (dateCol == "年月") DateTime.TryParse(r["年月"].ToString() + "-01", out rowDate);

                                        double price = GetPriceForDate(cfg.PriceCategory, rowDate);
                                        totalCost += (qty * price);
                                    }
                                }
                            }
                        }
                    }
                } catch { }

                sectionTotalCost += totalCost;

                // 產生卡片
                Panel card = new Panel { Size = new Size(280, 130), BackColor = Color.WhiteSmoke, Margin = new Padding(10) };
                card.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, card.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                
                Label lTitle = new Label { Text = cfg.DisplayName, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.FromArgb(60,60,60), Location = new Point(15,10), AutoSize = true };
                
                Label lQty = new Label { Text = $"總耗用：{totalQty:N2}", Font = new Font("Microsoft JhengHei UI", 11F), ForeColor = Color.DimGray, Location = new Point(15, 45), AutoSize = true };
                
                double avgPrice = totalQty > 0 ? totalCost / totalQty : 0;
                Label lAvg = new Label { Text = $"均單價：${avgPrice:N2}", Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.DimGray, Location = new Point(15, 70), AutoSize = true };

                Label lCost = new Label { Text = $"$ {totalCost:N0}", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), ForeColor = themeColor, Location = new Point(10, 90), AutoSize = true };

                card.Controls.AddRange(new Control[] { lTitle, lQty, lAvg, lCost });
                flp.Controls.Add(card);
            }

            // 總計卡片
            Panel totalCard = new Panel { Size = new Size(280, 130), BackColor = Color.LightYellow, Margin = new Padding(10) };
            totalCard.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, totalCard.ClientRectangle, Color.Orange, ButtonBorderStyle.Solid);
            Label lTotTitle = new Label { Text = "【區塊總計成本】", Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.DarkRed, Location = new Point(15,15), AutoSize = true };
            Label lTotCost = new Label { Text = $"$ {sectionTotalCost:N0}", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.Crimson, Location = new Point(10, 65), AutoSize = true };
            totalCard.Controls.AddRange(new Control[] { lTotTitle, lTotCost });
            flp.Controls.Add(totalCard);
        }

        private double GetPriceForDate(string category, DateTime date)
        {
            var matchedPrices = _prices.Where(p => p.Category == category && date >= p.StartDate && date <= p.EndDate).ToList();
            if (matchedPrices.Count > 0) return matchedPrices.First().UnitPrice;
            
            // 若找不到該區間，嘗試找該類別最新的單價
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
        // 統計項目設定視窗
        // ==========================================
        private void OpenConfigManager(string sectionCode)
        {
            using (Form f = new Form { Text = "⚙️ 成本統計項目設定", Size = new Size(850, 500), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                var dbMap = App_DbConfig.GetDbMapCache();
                var uniqueCategories = _prices.Select(p => p.Category).Distinct().ToList();

                DataGridView dgv = new DataGridView { 
                    Dock = DockStyle.Fill, BackgroundColor = Color.WhiteSmoke, AllowUserToAddRows = true, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Microsoft JhengHei UI", 11F) 
                };
                
                dgv.Columns.Add("Id", "Id"); dgv.Columns["Id"].Visible = false;
                dgv.Columns.Add("DisplayName", "顯示名稱 (例: PAC藥劑費)");
                
                DataGridViewComboBoxColumn cboDb = new DataGridViewComboBoxColumn { Name="DbName", HeaderText="來源資料庫" };
                foreach (var db in dbMap.Keys) cboDb.Items.Add(db);
                dgv.Columns.Add(cboDb);

                dgv.Columns.Add("TableName", "資料表 (手動填入英文名)");
                dgv.Columns.Add("ColName", "統計欄位名");
                
                DataGridViewComboBoxColumn cboPrice = new DataGridViewComboBoxColumn { Name="PriceCategory", HeaderText="綁定計價類別" };
                foreach (var cat in uniqueCategories) cboPrice.Items.Add(cat);
                dgv.Columns.Add(cboPrice);

                foreach(var c in _configs.Where(x => x.Section == sectionCode)) {
                    int rIdx = dgv.Rows.Add();
                    dgv.Rows[rIdx].Cells["Id"].Value = c.Id;
                    dgv.Rows[rIdx].Cells["DisplayName"].Value = c.DisplayName;
                    if (cboDb.Items.Contains(c.DbName)) dgv.Rows[rIdx].Cells["DbName"].Value = c.DbName;
                    dgv.Rows[rIdx].Cells["TableName"].Value = c.TableName;
                    dgv.Rows[rIdx].Cells["ColName"].Value = c.ColName;
                    if (cboPrice.Items.Contains(c.PriceCategory)) dgv.Rows[rIdx].Cells["PriceCategory"].Value = c.PriceCategory;
                }

                Button btnSave = new Button { Text = "💾 儲存設定", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                btnSave.Click += (s, e) => {
                    dgv.EndEdit();
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int)); dt.Columns.Add("Section", typeof(string)); dt.Columns.Add("DisplayName", typeof(string)); 
                    dt.Columns.Add("DbName", typeof(string)); dt.Columns.Add("TableName", typeof(string)); dt.Columns.Add("ColName", typeof(string)); dt.Columns.Add("PriceCategory", typeof(string));

                    foreach(DataGridViewRow r in dgv.Rows) {
                        if (r.IsNewRow) continue;
                        string disp = r.Cells["DisplayName"].Value?.ToString();
                        string db = r.Cells["DbName"].Value?.ToString();
                        string tb = r.Cells["TableName"].Value?.ToString();
                        string col = r.Cells["ColName"].Value?.ToString();
                        string priceCat = r.Cells["PriceCategory"].Value?.ToString();

                        if (!string.IsNullOrWhiteSpace(disp) && !string.IsNullOrWhiteSpace(db) && !string.IsNullOrWhiteSpace(tb) && !string.IsNullOrWhiteSpace(col)) {
                            DataRow dr = dt.NewRow();
                            if (r.Cells["Id"].Value != null && int.TryParse(r.Cells["Id"].Value.ToString(), out int id)) dr["Id"] = id;
                            dr["Section"] = sectionCode; dr["DisplayName"] = disp; dr["DbName"] = db; dr["TableName"] = tb; dr["ColName"] = col; dr["PriceCategory"] = priceCat ?? "";
                            dt.Rows.Add(dr);
                        }
                    }

                    // 刪除此 Section 舊資料
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"DELETE FROM {ConfigTable} WHERE Section='{sectionCode}'", conn)) cmd.ExecuteNonQuery();
                    }
                    
                    DataManager.BulkSaveTable(SysDbName, ConfigTable, dt);
                    LoadCache();
                    MessageBox.Show("設定儲存成功！", "成功");
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(dgv);
                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
        }

        // ==========================================
        // 輔助與導出
        // ==========================================
        private void ExportToPdf()
        {
            MessageBox.Show("水資源成本分析 PDF 匯出功能開發中，敬請期待。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
