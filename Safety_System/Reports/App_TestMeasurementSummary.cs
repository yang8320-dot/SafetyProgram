/// FILE: Safety_System/Reports/App_TestMeasurementSummary.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TestMeasurementSummary
    {
        private ComboBox _cboYear;
        private Button _btnSearch;
        private Button _btnPdf;
        private Button _btnSettings;
        private DataGridView _dgvResult;
        private DataTable _dtResult;
        private Panel _pnlGridContainer;

        // 設定檔路徑與快取
        private readonly string SettingsFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestSummarySettings.txt");
        private List<SummaryConfigItem> _configs = new List<SummaryConfigItem>();
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 目標資料表清單
        private readonly string[] _targetTables = { 
            "EnvMonitor", "WastewaterPeriodic", "DrinkingWater", "IndustrialZoneTest", 
            "SoilGasTest", "WastewaterSelfTest", "CoolingWaterVendor", "CoolingWaterSelf", 
            "TCLP", "WaterMeterCalibration", "OtherTests" 
        };

        // 定義自訂統計項目的資料結構
        private class SummaryConfigItem
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public string DateCol { get; set; }
            public string ItemCol { get; set; }
            public string PointCol { get; set; }
            public string ValueCol { get; set; }
            public string LimitCol { get; set; }
        }

        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public Control GetView()
        {
            _dbMap = App_DbConfig.GetDbMapCache();
            LoadSettings();

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            // ==========================================
            // 第一個框：資料選擇與操作列
            // ==========================================
            GroupBox box1 = new GroupBox { Text = "⚙️ 查詢條件與操作區", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0,0,0,20) };
            
            FlowLayoutPanel flpRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,5,0,10), WrapContents = false };
            
            _cboYear = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 13F), Margin = new Padding(0, 3, 0, 0) };
            int currentYear = DateTime.Today.Year;
            for (int i = currentYear - 10; i <= currentYear + 2; i++) {
                _cboYear.Items.Add(i.ToString());
            }
            _cboYear.SelectedItem = currentYear.ToString();

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(120, 38), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(15, 0, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += BtnSearch_Click;

            _btnSettings = new Button { Text = "⚙️ 統計設定", Size = new Size(130, 38), BackColor = Color.DimGray, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnSettings.FlatAppearance.BorderSize = 0;
            _btnSettings.Click += BtnSettings_Click;

            _btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(140, 38), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnPdf.FlatAppearance.BorderSize = 0;
            _btnPdf.Click += BtnPdf_Click;

            flpRow.Controls.AddRange(new Control[] {
                new Label { Text = "查詢年度:", AutoSize = true, Margin = new Padding(10, 10, 5, 0) }, 
                _cboYear,
                _btnSearch, _btnSettings, _btnPdf
            });

            box1.Controls.Add(flpRow);
            layout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 第二個框：預覽區
            // ==========================================
            GroupBox box2 = new GroupBox { Text = "📄 量測項目一覽表 (預覽區)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            _pnlGridContainer = new Panel { Dock = DockStyle.Top, Height = 600, BackColor = Color.White };
            _pnlGridContainer.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlGridContainer.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _dgvResult = new DataGridView {
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, 
                RowHeadersVisible = false, 
                Font = new Font("Microsoft JhengHei UI", 11F),
                Margin = new Padding(0), 
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.Single, 
                GridColor = Color.LightGray
            };
            
            _dgvResult.EnableHeadersVisualStyles = false;
            _dgvResult.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
            _dgvResult.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvResult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvResult.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            _dgvResult.ColumnHeadersHeight = 45;
            _dgvResult.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;

            _dtResult = new DataTable();
            _dtResult.Columns.Add("量測項目", typeof(string));
            _dtResult.Columns.Add("檢測點", typeof(string));
            _dtResult.Columns.Add("管制標準", typeof(string));
            for (int i = 1; i <= 12; i++) _dtResult.Columns.Add($"{i}月", typeof(string));
            _dtResult.Columns.Add("最大值", typeof(string));
            _dtResult.Columns.Add("最小值", typeof(string));
            _dtResult.Columns.Add("平均值", typeof(string));
            
            _dgvResult.DataSource = _dtResult;
            ApplyGridColumnWeights();

            _dgvResult.Resize += (s, e) => ApplyGridColumnWeights();

            _pnlGridContainer.Controls.Add(_dgvResult);
            box2.Controls.Add(_pnlGridContainer);
            layout.Controls.Add(box2, 0, 1);

            mainScrollPanel.Controls.Add(layout);
            
            if (_configs.Count > 0) BtnSearch_Click(null, null);

            return mainScrollPanel;
        }

        private void ApplyGridColumnWeights()
        {
            if (_dgvResult == null || _dgvResult.Columns.Count == 0) return;
            int totalGridWidth = _dgvResult.ClientSize.Width;
            if (totalGridWidth < 500) totalGridWidth = 1400; 

            float unitWidth = totalGridWidth / 100f;

            foreach (DataGridViewColumn col in _dgvResult.Columns) { col.MinimumWidth = 20; }

            _dgvResult.Columns["量測項目"].Width = (int)(unitWidth * 11f);
            _dgvResult.Columns["檢測點"].Width = (int)(unitWidth * 7f);
            _dgvResult.Columns["管制標準"].Width = (int)(unitWidth * 7f);

            for (int m = 1; m <= 12; m++) {
                _dgvResult.Columns[$"{m}月"].Width = (int)(unitWidth * 5f);
            }

            _dgvResult.Columns["最大值"].Width = (int)(unitWidth * 5f);
            _dgvResult.Columns["最小值"].Width = (int)(unitWidth * 5f);
            _dgvResult.Columns["平均值"].Width = (int)(unitWidth * 5f);
        }

        // =========================================================
        // 萬能日期解析引擎 (處理 年/年月/年月日 與 民國年)
        // =========================================================
        private (string Year, int Month) ParseDateFlexible(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return ("", 0);
            dateStr = dateStr.Trim().Replace("/", "-");

            Regex twRegex = new Regex(@"^(?<year>\d{2,3})(?:-(?<month>\d{1,2}))?(?:-(?<day>\d{1,2}))?(?:\s+.*)?$");
            Match matchTw = twRegex.Match(dateStr);

            string yStr = "";
            int mInt = 0;

            if (matchTw.Success)
            {
                if (int.TryParse(matchTw.Groups["year"].Value, out int twYear))
                {
                    int finalYear = twYear < 200 ? twYear + 1911 : twYear;
                    yStr = finalYear.ToString();

                    if (matchTw.Groups["month"].Success) {
                        int.TryParse(matchTw.Groups["month"].Value, out mInt);
                    }
                    return (yStr, mInt);
                }
            }

            // 備用解析 (純西元年等)
            if (DateTime.TryParse(dateStr, out DateTime result)) {
                return (result.Year.ToString(), result.Month);
            }

            return ("", 0);
        }

        // =========================================================
        // 🟢 智慧欄位容錯補救引擎 (Intelligent Fallback Engine)
        // =========================================================
        private string GetFallbackDateCol(DataTable dt) => dt.Columns.Contains("日期") ? "日期" : (dt.Columns.Contains("年月") ? "年月" : (dt.Columns.Contains("年度") ? "年度" : ""));
        private string GetFallbackItemCol(DataTable dt) => new[] { "檢測項目", "項目", "名稱", "設備名稱", "量測項目" }.FirstOrDefault(c => dt.Columns.Contains(c)) ?? "";
        private string GetFallbackPointCol(DataTable dt) => new[] { "檢測點", "點位", "SEG編號", "水錶名稱", "位置" }.FirstOrDefault(c => dt.Columns.Contains(c)) ?? "";
        private string GetFallbackValueCol(DataTable dt) => new[] { "檢測數據", "現場流量計讀值", "數值", "結果" }.FirstOrDefault(c => dt.Columns.Contains(c)) ?? "";
        private string GetFallbackLimitCol(DataTable dt) => new[] { "管制值", "管制標準", "標準" }.FirstOrDefault(c => dt.Columns.Contains(c)) ?? "";

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (_cboYear.SelectedItem == null) return;
            string targetYear = _cboYear.SelectedItem.ToString();
            
            if (_configs.Count == 0) {
                MessageBox.Show("目前尚未設定任何資料來源，請點擊【統計設定】新增來源欄位！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try 
            {
                _dtResult.Rows.Clear(); 
                // 🟢 字典結構修改：保存 (純數值 Num, 原始字串 Raw)
                var aggregatedData = new Dictionary<string, Dictionary<int, List<(double Num, string Raw)>>>();

                foreach (var cfg in _configs) 
                {
                    DataTable dt = null;
                    try {
                        dt = DataManager.GetTableData(cfg.DbName, cfg.TableName, "", "", "");
                    } catch { continue; }

                    if (dt == null || dt.Rows.Count == 0) continue;

                    string actualDateCol = !string.IsNullOrEmpty(cfg.DateCol) && dt.Columns.Contains(cfg.DateCol) ? cfg.DateCol : GetFallbackDateCol(dt);
                    string actualItemCol = !string.IsNullOrEmpty(cfg.ItemCol) && dt.Columns.Contains(cfg.ItemCol) ? cfg.ItemCol : GetFallbackItemCol(dt);
                    string actualPointCol = !string.IsNullOrEmpty(cfg.PointCol) && dt.Columns.Contains(cfg.PointCol) ? cfg.PointCol : GetFallbackPointCol(dt);
                    string actualValueCol = !string.IsNullOrEmpty(cfg.ValueCol) && dt.Columns.Contains(cfg.ValueCol) ? cfg.ValueCol : GetFallbackValueCol(dt);
                    string actualLimitCol = !string.IsNullOrEmpty(cfg.LimitCol) && dt.Columns.Contains(cfg.LimitCol) ? cfg.LimitCol : GetFallbackLimitCol(dt);

                    if (string.IsNullOrEmpty(actualDateCol) || string.IsNullOrEmpty(actualItemCol) || string.IsNullOrEmpty(actualValueCol)) continue;

                    foreach (DataRow r in dt.Rows) 
                    {
                        if (r.RowState == DataRowState.Deleted) continue;

                        string rawDate = r[actualDateCol]?.ToString() ?? "";
                        var parsedDate = ParseDateFlexible(rawDate);
                        
                        if (parsedDate.Year != targetYear) continue; 

                        string item = r[actualItemCol]?.ToString()?.Trim() ?? "";
                        string point = !string.IsNullOrEmpty(actualPointCol) ? r[actualPointCol]?.ToString()?.Trim() ?? "" : "";
                        string limit = !string.IsNullOrEmpty(actualLimitCol) ? r[actualLimitCol]?.ToString()?.Trim() ?? "" : "";
                        string valStr = r[actualValueCol]?.ToString()?.Replace(",", "").Trim() ?? "";

                        if (string.IsNullOrEmpty(item) || string.IsNullOrEmpty(valStr)) continue;

                        // ========================================================
                        // 🟢 智慧數值萃取引擎 (保留原始 % 或 < 顯示，但提取純數字運算)
                        // ========================================================
                        double numericVal = 0;
                        bool isNumericValid = false;

                        // 判斷是否為「未檢出」或 N.D.，數學上視為 0
                        if (valStr.ToUpper().Contains("N.D") || valStr.Contains("未檢出")) {
                            numericVal = 0;
                            isNumericValid = true;
                        } else {
                            // 擷取所有可能是小數點或負數的連續數字
                            Match mNum = Regex.Match(valStr, @"[-+]?[0-9]*\.?[0-9]+");
                            if (mNum.Success && double.TryParse(mNum.Value, out double parsedVal)) {
                                numericVal = parsedVal;
                                isNumericValid = true;
                            }
                        }

                        if (isNumericValid) 
                        {
                            string key = $"{item}|{point}|{limit}";

                            if (!aggregatedData.ContainsKey(key)) {
                                aggregatedData[key] = new Dictionary<int, List<(double Num, string Raw)>>();
                                for (int i = 0; i <= 12; i++) aggregatedData[key][i] = new List<(double Num, string Raw)>();
                            }

                            int mIdx = (parsedDate.Month >= 1 && parsedDate.Month <= 12) ? parsedDate.Month : 0;
                            aggregatedData[key][mIdx].Add((numericVal, valStr)); // 將數值與原始字面一同保存
                        }
                    }
                }

                foreach (var kvp in aggregatedData) 
                {
                    var parts = kvp.Key.Split('|');
                    string item = parts[0];
                    string point = parts[1];
                    string limit = parts[2];

                    DataRow row = _dtResult.NewRow();
                    row["量測項目"] = item;
                    row["檢測點"] = point;
                    row["管制標準"] = limit;

                    List<double> allValuesForYear = new List<double>();
                    
                    // 將沒有月份的數據先加入總池 (只取純數值參與運算)
                    allValuesForYear.AddRange(kvp.Value[0].Select(x => x.Num));

                    for (int m = 1; m <= 12; m++) 
                    {
                        var mValues = kvp.Value[m];
                        if (mValues.Count > 0) 
                        {
                            // 🟢 畫面顯示：輸出原始字面字串 (包含 %, <, >, N.D. 等)
                            row[$"{m}月"] = string.Join("、", mValues.Select(v => v.Raw));
                            // 🟢 數學運算：將純數字加入平均/最大/最小池中
                            allValuesForYear.AddRange(mValues.Select(x => x.Num));
                        } 
                        else 
                        {
                            row[$"{m}月"] = "";
                        }
                    }

                    // 結算整年的最大、最小、平均
                    if (allValuesForYear.Count > 0) {
                        row["最大值"] = allValuesForYear.Max().ToString("0.##");
                        row["最小值"] = allValuesForYear.Min().ToString("0.##");
                        row["平均值"] = allValuesForYear.Average().ToString("0.##");
                    } else {
                        row["最大值"] = ""; row["最小值"] = ""; row["平均值"] = "";
                    }

                    _dtResult.Rows.Add(row);
                }

                _dtResult.DefaultView.Sort = "量測項目 ASC, 檢測點 ASC";
                _dtResult = _dtResult.DefaultView.ToTable();

                _dgvResult.DataSource = _dtResult;
                ApplyGridColumnWeights();
                _dgvResult.ClearSelection();
            } 
            catch (Exception ex) 
            {
                MessageBox.Show("查詢失敗：" + ex.Message, "錯誤");
            } 
            finally 
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (_dgvResult == null || _dgvResult.Rows.Count == 0) {
                MessageBox.Show("目前沒有數據可供導出，請先執行查詢。"); return;
            }

            string targetYear = _cboYear.SelectedItem.ToString();

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = $"環境量測項目績效一覽表_{targetYear}年度_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true; // 橫式
                    pd.DefaultPageSettings.Margins = new Margins(40, 40, 50, 50);

                    int currentRowIndex = 0;
                    int pageNumber = 1;
                    
                    // 先計算總頁數
                    int totalPages = 1;
                    float simH = 827f - 100f; // A4橫向高度扣掉上下邊界
                    float simStartY = 50f + 40f + 40f + 35f + 40f + 35f; // 標題高度
                    float simCurrentY = simStartY;
                    float simBottomLimit = simH - 30f; 

                    for (int i = 0; i < _dgvResult.Rows.Count; i++) {
                        float rowH = 35f; 
                        if (simCurrentY + rowH > simBottomLimit) {
                            totalPages++;
                            simCurrentY = simStartY + rowH;
                        } else {
                            simCurrentY += rowH;
                        }
                    }

                    pd.PrintPage += (s, ev) => {
                        Graphics g = ev.Graphics;
                        float x = ev.MarginBounds.Left;
                        float y = ev.MarginBounds.Top;
                        float w = ev.MarginBounds.Width;

                        Font fMainTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                        Font fSubTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                        Font fSign = new Font("Microsoft JhengHei UI", 12F);
                        
                        Font fGridTitle = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
                        Font fGridHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                        Font fGridBody = new Font("Microsoft JhengHei UI", 9F);
                        Font fFooter = new Font("Microsoft JhengHei UI", 10F);

                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        // 1. 大標題與簽核 (每頁都有)
                        g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fMainTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                        y += 40;
                        g.DrawString("環境量測項目績效一覽表", fSubTitle, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                        y += 40;
                        string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
                        g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                        y += 40;

                        // 2. 雙層表頭繪製
                        float rowH = 35f;
                        float[] colWidths = new float[_dgvResult.Columns.Count];
                        
                        float unitWidth = w / 100f;
                        
                        colWidths[0] = unitWidth * 11f; // 量測項目 11%
                        colWidths[1] = unitWidth * 7f;  // 檢測點   7%
                        colWidths[2] = unitWidth * 7f;  // 管制標準 7%
                        
                        for (int i = 3; i <= 17; i++) {
                            colWidths[i] = unitWidth * 5f; // 1~12月 + 最大 + 最小 + 平均，共15格，每格5%
                        }

                        // 第一層表頭
                        float currX = x;
                        for (int i = 0; i <= 2; i++) {
                            RectangleF rHead = new RectangleF(currX, y, colWidths[i], rowH * 2);
                            g.FillRectangle(Brushes.LightGray, rHead);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rHead));
                            g.DrawString(_dgvResult.Columns[i].HeaderText, fGridTitle, Brushes.Black, rHead, sfCenter);
                            currX += colWidths[i];
                        }
                        
                        // 右側群組標題「XX 年度」
                        float monthGroupWidth = unitWidth * 75f; 
                        RectangleF rYearTitle = new RectangleF(currX, y, monthGroupWidth, rowH);
                        g.FillRectangle(Brushes.LightGray, rYearTitle);
                        g.DrawRectangle(Pens.Black, Rectangle.Round(rYearTitle));
                        g.DrawString($"{targetYear} 年度", fGridTitle, Brushes.Black, rYearTitle, sfCenter);

                        // 第二層表頭 (1~12月, 最大, 最小, 平均)
                        float subY = y + rowH;
                        float subX = currX;
                        for (int i = 3; i < _dgvResult.Columns.Count; i++) {
                            RectangleF rSubHead = new RectangleF(subX, subY, colWidths[i], rowH);
                            g.FillRectangle(Brushes.LightGray, rSubHead);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rSubHead));
                            g.DrawString(_dgvResult.Columns[i].HeaderText, fGridHead, Brushes.Black, rSubHead, sfCenter);
                            subX += colWidths[i];
                        }
                        
                        y += rowH * 2;

                        // 3. 資料清單
                        while (currentRowIndex < _dgvResult.Rows.Count) 
                        {
                            DataGridViewRow row = _dgvResult.Rows[currentRowIndex];
                            
                            if (y + rowH > ev.MarginBounds.Bottom - 30) {
                                g.DrawString("8-ES-B11-02 環境量測項目績效一覽表", fFooter, Brushes.Black, x, ev.MarginBounds.Bottom - 15);
                                g.DrawString($"{pageNumber} / {totalPages}", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom - 15, w, 20), sfCenter);
                                pageNumber++;
                                ev.HasMorePages = true;
                                return;
                            }

                            currX = x;
                            for (int i = 0; i < _dgvResult.Columns.Count; i++) {
                                RectangleF rCell = new RectangleF(currX, y, colWidths[i], rowH);
                                g.DrawRectangle(Pens.Black, Rectangle.Round(rCell));
                                string val = row.Cells[i].Value?.ToString() ?? "";
                                g.DrawString(val, fGridBody, Brushes.Black, rCell, sfCenter);
                                currX += colWidths[i];
                            }
                            y += rowH;
                            currentRowIndex++;
                        }

                        // 4. 補空白列填滿畫面
                        while (y + rowH <= ev.MarginBounds.Bottom - 30)
                        {
                            currX = x;
                            for (int i = 0; i < _dgvResult.Columns.Count; i++) {
                                RectangleF rCell = new RectangleF(currX, y, colWidths[i], rowH);
                                g.DrawRectangle(Pens.Black, Rectangle.Round(rCell));
                                currX += colWidths[i];
                            }
                            y += rowH;
                        }

                        // 5. 底部代碼與頁碼
                        g.DrawString("8-ES-B11-02 環境量測項目績效一覽表", fFooter, Brushes.Black, x, ev.MarginBounds.Bottom - 15);
                        g.DrawString($"{pageNumber} / {totalPages}", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom - 15, w, 20), sfCenter);

                        ev.HasMorePages = false;
                    };

                    try {
                        pd.Print();
                        MessageBox.Show("PDF 導出成功！\n匯出內容已完全同步目前的視窗排序狀態。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 導出失敗：" + ex.Message, "錯誤");
                    }
                }
            }
        }

        // =========================================================
        // 設定檔管理與動態設定視窗 (加入 Memory Cache 以解決卡頓)
        // =========================================================
        private void LoadSettings()
        {
            _configs.Clear();
            if (System.IO.File.Exists(SettingsFile))
            {
                try {
                    foreach (var line in System.IO.File.ReadAllLines(SettingsFile, System.Text.Encoding.UTF8)) {
                        var parts = line.Split('|');
                        if (parts.Length == 7) {
                            _configs.Add(new SummaryConfigItem { 
                                DbName = parts[0], TableName = parts[1], 
                                DateCol = parts[2], ItemCol = parts[3], 
                                PointCol = parts[4], ValueCol = parts[5], LimitCol = parts[6]
                            });
                        }
                    }
                } catch { }
            }
            
            if (_configs.Count == 0) {
                foreach (var tb in _targetTables) {
                    _configs.Add(new SummaryConfigItem { DbName = "TestData", TableName = tb, DateCol = "日期", ItemCol = "檢測項目", PointCol = "檢測點", ValueCol = "檢測數據", LimitCol = "管制值" });
                }
            }
        }

        private void SaveSettings()
        {
            try {
                var lines = _configs.Select(c => $"{c.DbName}|{c.TableName}|{c.DateCol}|{c.ItemCol}|{c.PointCol}|{c.ValueCol}|{c.LimitCol}").ToArray();
                System.IO.File.WriteAllLines(SettingsFile, lines, System.Text.Encoding.UTF8);
            } catch { }
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "⚙️ 讀取資料來源設定 (定義要合併統計的資料表)", Size = new Size(1300, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Panel pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.WhiteSmoke };
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 8, RowCount = 15, Padding = new Padding(10) };
                
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50F));  // 刪除按鈕
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F)); // 庫
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F)); // 表
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));   // 日期
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));   // 項目
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));   // 檢測點
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));   // 數據
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));   // 管制值

                string[] headers = { "", "來源資料庫", "來源資料表", "對應[日期]欄位", "對應[量測項目]欄", "對應[檢測點]欄", "對應[檢測數據]欄", "對應[管制值]欄" };
                for (int i = 0; i < 8; i++) tlp.Controls.Add(new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill }, i, 0);

                var editingConfigs = new List<SummaryConfigItem>(_configs);

                Dictionary<string, List<string>> _columnCache = new Dictionary<string, List<string>>();

                Action renderRows = null;
                renderRows = () => {
                    tlp.SuspendLayout();
                    pnlScroll.SuspendLayout();

                    while (tlp.Controls.Count > 8) tlp.Controls.RemoveAt(8);
                    tlp.RowCount = editingConfigs.Count + 2;

                    for (int i = 0; i < editingConfigs.Count; i++) {
                        int currentIndex = i;
                        var conf = editingConfigs[i];

                        Button btnDel = new Button { Text = "❌", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        btnDel.FlatAppearance.BorderSize = 0;
                        btnDel.Click += (s, ev) => { editingConfigs.RemoveAt(currentIndex); renderRows(); };

                        ComboBox cbDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbTb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbDate = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbItem = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbPoint = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbVal = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                        ComboBox cbLimit = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };

                        foreach (var kvp in _dbMap) cbDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

                        Action<string, string> loadCols = (dbEnName, tbEnName) => {
                            cbDate.Items.Clear(); cbItem.Items.Clear(); cbPoint.Items.Clear(); cbVal.Items.Clear(); cbLimit.Items.Clear();
                            cbDate.Items.Add(""); cbItem.Items.Add(""); cbPoint.Items.Add(""); cbVal.Items.Add(""); cbLimit.Items.Add("");
                            
                            if (!string.IsNullOrEmpty(tbEnName) && !string.IsNullOrEmpty(dbEnName)) {
                                string cacheKey = $"{dbEnName}_{tbEnName}";
                                List<string> cols;
                                
                                if (_columnCache.ContainsKey(cacheKey)) {
                                    cols = _columnCache[cacheKey]; 
                                } else {
                                    cols = DataManager.GetColumnNames(dbEnName, tbEnName);
                                    _columnCache[cacheKey] = cols; 
                                }

                                foreach(var c in cols) {
                                    if (c != "Id") {
                                        cbDate.Items.Add(c); cbItem.Items.Add(c); cbPoint.Items.Add(c); cbVal.Items.Add(c); cbLimit.Items.Add(c);
                                    }
                                }
                            }
                        };

                        cbDb.SelectedIndexChanged += (s, ev) => {
                            cbTb.Items.Clear();
                            if (cbDb.SelectedItem != null) {
                                conf.DbName = ((ItemMap)cbDb.SelectedItem).EnName;
                                foreach(var tb in _dbMap[conf.DbName].Tables) cbTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                            }
                        };

                        cbTb.SelectedIndexChanged += (s, ev) => {
                            if (cbTb.SelectedItem != null && cbDb.SelectedItem != null) {
                                conf.TableName = ((ItemMap)cbTb.SelectedItem).EnName;
                                string dName = ((ItemMap)cbDb.SelectedItem).EnName;
                                loadCols(dName, conf.TableName);
                            }
                        };

                        cbDate.SelectedIndexChanged += (s, ev) => { conf.DateCol = cbDate.Text; };
                        cbItem.SelectedIndexChanged += (s, ev) => { conf.ItemCol = cbItem.Text; };
                        cbPoint.SelectedIndexChanged += (s, ev) => { conf.PointCol = cbPoint.Text; };
                        cbVal.SelectedIndexChanged += (s, ev) => { conf.ValueCol = cbVal.Text; };
                        cbLimit.SelectedIndexChanged += (s, ev) => { conf.LimitCol = cbLimit.Text; };

                        foreach (ItemMap im in cbDb.Items) if (im.EnName == conf.DbName) { cbDb.SelectedItem = im; break; }
                        if (cbDb.SelectedItem != null) {
                            foreach (ItemMap im in cbTb.Items) if (im.EnName == conf.TableName) { cbTb.SelectedItem = im; break; }
                            if (cbTb.SelectedItem != null) loadCols(conf.DbName, conf.TableName);
                        }
                        
                        if (!string.IsNullOrEmpty(conf.DateCol) && cbDate.Items.Contains(conf.DateCol)) cbDate.SelectedItem = conf.DateCol;
                        if (!string.IsNullOrEmpty(conf.ItemCol) && cbItem.Items.Contains(conf.ItemCol)) cbItem.SelectedItem = conf.ItemCol;
                        if (!string.IsNullOrEmpty(conf.PointCol) && cbPoint.Items.Contains(conf.PointCol)) cbPoint.SelectedItem = conf.PointCol;
                        if (!string.IsNullOrEmpty(conf.ValueCol) && cbVal.Items.Contains(conf.ValueCol)) cbVal.SelectedItem = conf.ValueCol;
                        if (!string.IsNullOrEmpty(conf.LimitCol) && cbLimit.Items.Contains(conf.LimitCol)) cbLimit.SelectedItem = conf.LimitCol;

                        tlp.Controls.Add(btnDel, 0, i + 1);
                        tlp.Controls.Add(cbDb, 1, i + 1);
                        tlp.Controls.Add(cbTb, 2, i + 1);
                        tlp.Controls.Add(cbDate, 3, i + 1);
                        tlp.Controls.Add(cbItem, 4, i + 1);
                        tlp.Controls.Add(cbPoint, 5, i + 1);
                        tlp.Controls.Add(cbVal, 6, i + 1);
                        tlp.Controls.Add(cbLimit, 7, i + 1);
                    }

                    Button btnAdd = new Button { Text = "➕ 新增來源", Dock = DockStyle.Fill, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    btnAdd.Click += (s, ev) => { editingConfigs.Add(new SummaryConfigItem()); renderRows(); };
                    
                    tlp.Controls.Add(btnAdd, 1, editingConfigs.Count + 1);
                    tlp.SetColumnSpan(btnAdd, 7);

                    pnlScroll.ResumeLayout(false);
                    tlp.ResumeLayout(true);
                };

                renderRows();
                pnlScroll.Controls.Add(tlp);

                Button btnSave = new Button { Text = "💾 儲存設定並重新載入", Dock = DockStyle.Bottom, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSave.FlatAppearance.BorderSize = 0;
                
                btnSave.Click += (senderObj, evnt) => {
                    _configs = editingConfigs;
                    SaveSettings();
                    BtnSearch_Click(null, null);
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(pnlScroll);
                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
        }
    }
}
