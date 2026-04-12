/// FILE: Safety_System/App_WaterDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        private DateTimePicker _dtpStart, _dtpEnd;
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
            Panel box1 = new Panel { Dock = DockStyle.Fill, Height = 80, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            box1.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, box1.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            Label lblTitle = new Label { Text = "💧 水資源綜合數據看板", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Location = new Point(20, 22) };
            
            FlowLayoutPanel flpControls = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 22, 20, 0) };
            
            _dtpStart = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 130, Font = new Font("Microsoft JhengHei UI", 12F), Value = DateTime.Today.AddMonths(-1) };
            _dtpEnd = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 130, Font = new Font("Microsoft JhengHei UI", 12F), Value = DateTime.Today };
            
            Button btnSearch = new Button { Text = "🔍 查詢統計", Size = new Size(120, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSearch.Click += (s, e) => LoadAllData();
            
            Button btnPdf = new Button { Text = "📄 轉存 PDF", Size = new Size(120, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _dtpStart, new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _dtpEnd,
                new Panel { Width = 10, Height = 1 }, btnSearch, btnPdf
            });

            box1.Controls.Add(lblTitle);
            box1.Controls.Add(flpControls);
            mainLayout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 建立 3 個主要數據區塊 (大框 2, 3, 4)
            // ==========================================
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 水資源數據統計", Color.Teal, out _pnlBox2Data1, out _pnlBox2Data2, out _pnlBox2Data3, out _pnlBox2Data4), 0, 1);
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 回收水統計", Color.ForestGreen, out _pnlBox3Data1, out _pnlBox3Data2, out _pnlBox3Data3, out _pnlBox3Data4), 0, 2);
            mainLayout.Controls.Add(BuildNineGridBox("台灣玻璃彰濱廠 - 藥劑數據統計", Color.Sienna, out _pnlBox4Data1, out _pnlBox4Data2, out _pnlBox4Data3, out _pnlBox4Data4), 0, 3);

            _mainScrollPanel.Controls.Add(mainLayout);
            LoadAllData(); // 初始化載入
            return _mainScrollPanel;
        }

        // ==========================================
        // UI 產生器：建立九宮格等距大框
        // ==========================================
        private Panel BuildNineGridBox(string mainTitle, Color headerColor, out Panel d1, out Panel d2, out Panel d3, out Panel d4)
        {
            Panel outer = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            outer.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, outer.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel grid = new TableLayoutPanel { 
                Dock = DockStyle.Top, AutoSize = true, 
                ColumnCount = 4, RowCount = 3, 
                Padding = new Padding(10) 
            };
            
            // 四個欄位等距 25%
            for (int i = 0; i < 4; i++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); // 第一排 標題
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // 第二排 副標題
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 第三排 數據區

            // 第一排：全版面標題
            Label lblMainTitle = new Label { Text = mainTitle, Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };
            grid.Controls.Add(lblMainTitle, 0, 0);
            grid.SetColumnSpan(lblMainTitle, 4);

            // 第二排：四個小標題
            string[] subTitles;
            if (mainTitle.Contains("回收水")) {
                subTitles = new[] { "區間回收水量統計", "區間去年同期回收水量數據", "區間前年同期回收水量數據", "區間差異分析" };
            } else {
                subTitles = new[] { "區間用量統計", "區間去年同期數據", "區間前年同期數據", "區間差異分析" };
            }

            for (int i = 0; i < 4; i++) {
                Label l = new Label { Text = subTitles[i], Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.White, BackColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
                grid.Controls.Add(l, i, 1);
            }

            // 第三排：四個資料存放 Panel (透過 out 參數回傳給主程式控制)
            d1 = CreateDataPanel(); d2 = CreateDataPanel(); d3 = CreateDataPanel(); d4 = CreateDataPanel();
            grid.Controls.Add(d1, 0, 2); grid.Controls.Add(d2, 1, 2); grid.Controls.Add(d3, 2, 2); grid.Controls.Add(d4, 3, 2);

            outer.Controls.Add(grid);
            return outer;
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

            string dS = _dtpStart.Value.ToString("yyyy-MM-dd");
            string dE = _dtpEnd.Value.ToString("yyyy-MM-dd");
            
            string dS_LY = _dtpStart.Value.AddYears(-1).ToString("yyyy-MM-dd");
            string dE_LY = _dtpEnd.Value.AddYears(-1).ToString("yyyy-MM-dd");
            
            string dS_L2Y = _dtpStart.Value.AddYears(-2).ToString("yyyy-MM-dd");
            string dE_L2Y = _dtpEnd.Value.AddYears(-2).ToString("yyyy-MM-dd");

            // --- 大框 2：水資源數據統計 (WaterMeterReadings + WaterUsageDaily) ---
            var sumBox2_Curr = GetSumsEndingWith(dS, dE, "WaterMeterReadings", "WaterUsageDaily");
            var sumBox2_LY = GetSumsEndingWith(dS_LY, dE_LY, "WaterMeterReadings", "WaterUsageDaily");
            var sumBox2_L2Y = GetSumsEndingWith(dS_L2Y, dE_L2Y, "WaterMeterReadings", "WaterUsageDaily");
            
            FillDataPanels(_pnlBox2Data1, _pnlBox2Data2, _pnlBox2Data3, _pnlBox2Data4, sumBox2_Curr, sumBox2_LY, sumBox2_L2Y);

            // --- 大框 3：回收水統計 (專屬邏輯 WaterMeterReadings) ---
            var recycleCurr = CalculateRecycleStats(dS, dE);
            var recycleLY = CalculateRecycleStats(dS_LY, dE_LY);
            var recycleL2Y = CalculateRecycleStats(dS_L2Y, dE_L2Y);
            
            FillDataPanels(_pnlBox3Data1, _pnlBox3Data2, _pnlBox3Data3, _pnlBox3Data4, recycleCurr, recycleLY, recycleL2Y, true);

            // --- 大框 4：藥劑數據統計 (WaterChemicals) ---
            var sumBox4_Curr = GetSumsEndingWith(dS, dE, "WaterChemicals");
            var sumBox4_LY = GetSumsEndingWith(dS_LY, dE_LY, "WaterChemicals");
            var sumBox4_L2Y = GetSumsEndingWith(dS_L2Y, dE_L2Y, "WaterChemicals");

            FillDataPanels(_pnlBox4Data1, _pnlBox4Data2, _pnlBox4Data3, _pnlBox4Data4, sumBox4_Curr, sumBox4_LY, sumBox4_L2Y);

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        // ==========================================
        // 輔助方法：抓取多個資料表所有 "日統計" 結尾的加總
        // ==========================================
        private Dictionary<string, double> GetSumsEndingWith(string start, string end, params string[] tableNames)
        {
            var results = new Dictionary<string, double>();

            foreach (string tbl in tableNames)
            {
                DataTable dt = null;
                try { dt = DataManager.GetTableData(DbName, tbl, "日期", start, end); } catch { continue; }
                if (dt == null) continue;

                var targetCols = dt.Columns.Cast<DataColumn>().Where(c => c.ColumnName.EndsWith("日統計")).Select(c => c.ColumnName).ToList();

                foreach (DataRow r in dt.Rows) {
                    foreach (string col in targetCols) {
                        string cleanName = col.Replace("日統計", ""); // 畫面上移除 "日統計" 字眼比較好看
                        if (!results.ContainsKey(cleanName)) results[cleanName] = 0;
                        if (double.TryParse(r[col]?.ToString().Replace(",", ""), out double v)) {
                            results[cleanName] += v;
                        }
                    }
                }
            }
            return results;
        }

        // ==========================================
        // 輔助方法：計算大框 3 專屬回收率邏輯
        // ==========================================
        private Dictionary<string, double> CalculateRecycleStats(string start, string end)
        {
            var dict = new Dictionary<string, double> {
                { "廢水處理量", 0 }, { "回收水雙介質A", 0 }, { "回收水雙介質B", 0 }, { "總回收量", 0 }, { "回收率(%)", 0 }
            };

            DataTable dt = null;
            try { dt = DataManager.GetTableData(DbName, "WaterMeterReadings", "日期", start, end); } catch { return dict; }
            if (dt == null) return dict;

            foreach (DataRow r in dt.Rows) {
                if (dt.Columns.Contains("廢水處理量日統計") && double.TryParse(r["廢水處理量日統計"]?.ToString().Replace(",", ""), out double w)) dict["廢水處理量"] += w;
                if (dt.Columns.Contains("回收水雙介質A日統計") && double.TryParse(r["回收水雙介質A日統計"]?.ToString().Replace(",", ""), out double a)) dict["回收水雙介質A"] += a;
                if (dt.Columns.Contains("回收水雙介質B日統計") && double.TryParse(r["回收水雙介質B日統計"]?.ToString().Replace(",", ""), out double b)) dict["回收水雙介質B"] += b;
            }

            dict["總回收量"] = dict["回收水雙介質A"] + dict["回收水雙介質B"];
            if (dict["廢水處理量"] > 0) {
                dict["回收率(%)"] = (dict["總回收量"] / dict["廢水處理量"]) * 100;
            }

            return dict;
        }

        // ==========================================
        // 輔助方法：渲染四個區塊的 Label，包含差異分析 (YoY)
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

                // Panel 1: 當期
                p1.Controls.Add(CreateStatLabel(key, vCurr, isRecycleRate && key.Contains("%")));
                // Panel 2: 去年同期
                p2.Controls.Add(CreateStatLabel(key, vLy, isRecycleRate && key.Contains("%")));
                // Panel 3: 前年同期
                p3.Controls.Add(CreateStatLabel(key, vL2y, isRecycleRate && key.Contains("%")));

                // Panel 4: 差異分析 ((當期 - 去年) / 去年 * 100)
                string diffText = "無基期";
                Color diffColor = Color.DimGray;

                if (vLy > 0) {
                    double yoy = ((vCurr - vLy) / vLy) * 100;
                    
                    if (isRecycleRate && key.Contains("%")) {
                        // 如果本身就是百分比 (回收率)，差異直接相減即可 (絕對差異)
                        yoy = vCurr - vLy; 
                        diffText = (yoy > 0 ? "+" : "") + yoy.ToString("0.##") + " %";
                        diffColor = yoy > 0 ? Color.ForestGreen : Color.IndianRed; // 回收率高比較好
                    } else {
                        diffText = (yoy > 0 ? "+" : "") + yoy.ToString("0.##") + " %";
                        diffColor = yoy > 0 ? Color.IndianRed : Color.ForestGreen; // 用量低比較好
                    }
                } else if (vCurr > 0) {
                    diffText = "新數據";
                    diffColor = Color.SteelBlue;
                }

                Label lblDiff = new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
                p4.Controls.Add(lblDiff);
            }
        }

        private Label CreateStatLabel(string title, double value, bool isPercentage)
        {
            string format = isPercentage ? "0.##" : "N1";
            string suffix = isPercentage ? " %" : "";
            return new Label { 
                Text = $"{title}: {value.ToString(format)}{suffix}", 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                ForeColor = Color.FromArgb(45,45,45), 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 8) 
            };
        }

        // ==========================================
        // 匯出 PDF 功能
        // ==========================================
        private void BtnPdf_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "水資源數據統計_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        // 將 _mainScrollPanel 繪製成圖片 (這招能完美保留 UI 排版)
                        // 因為有 ScrollBar，需要暫時放大 Panel 確保拍到全貌
                        int originalHeight = _mainScrollPanel.Height;
                        _mainScrollPanel.Height = _mainScrollPanel.DisplayRectangle.Height; 
                        
                        Bitmap bmp = new Bitmap(_mainScrollPanel.Width, _mainScrollPanel.Height);
                        _mainScrollPanel.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                        
                        _mainScrollPanel.Height = originalHeight; // 恢復原狀

                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = false; // 直向列印
                        pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);

                        int currentY = 0; // 處理多頁裁切

                        pd.PrintPage += (s, ev) => {
                            Graphics g = ev.Graphics;
                            float scale = (float)ev.MarginBounds.Width / bmp.Width;
                            int sourceHeightFit = (int)(ev.MarginBounds.Height / scale);

                            Rectangle destRect = new Rectangle(ev.MarginBounds.Left, ev.MarginBounds.Top, ev.MarginBounds.Width, ev.MarginBounds.Height);
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

                        MessageBox.Show("PDF 匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
