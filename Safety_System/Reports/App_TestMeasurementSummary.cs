/// FILE: Safety_System/Reports/App_TestMeasurementSummary.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TestMeasurementSummary
    {
        private ComboBox _cboYear;
        private Button _btnSearch;
        private Button _btnPdf;
        private DataGridView _dgvResult;
        private DataTable _dtResult;

        // 需要被掃描的檢測數據表 (排除系統功能表與非檢測表)
        private readonly string[] _targetTables = { 
            "EnvMonitor", "WastewaterPeriodic", "DrinkingWater", "IndustrialZoneTest", 
            "SoilGasTest", "WastewaterSelfTest", "CoolingWaterVendor", "CoolingWaterSelf", 
            "TCLP", "OtherTests" 
        };

        public Control GetView()
        {
            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            // ==========================================
            // 第一個框：資料選擇與操作列
            // ==========================================
            GroupBox box1 = new GroupBox { Text = "⚙️ 查詢條件與操作區", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0,0,0,20) };
            
            FlowLayoutPanel flpRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,5,0,10), WrapContents = false };
            
            _cboYear = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            int currentYear = DateTime.Today.Year;
            for (int i = currentYear - 10; i <= currentYear + 2; i++) {
                _cboYear.Items.Add(i.ToString());
            }
            _cboYear.SelectedItem = currentYear.ToString();

            _btnSearch = new Button { Text = "🔍 查詢", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += BtnSearch_Click;

            _btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(140, 35), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            _btnPdf.FlatAppearance.BorderSize = 0;
            _btnPdf.Click += BtnPdf_Click;

            flpRow.Controls.AddRange(new Control[] {
                new Label { Text = "查詢年度:", AutoSize = true, Margin = new Padding(10,5,5,0) }, _cboYear,
                new Panel { Width=10, Height=1 }, _btnSearch, _btnPdf
            });

            box1.Controls.Add(flpRow);
            layout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 第二個框：預覽區
            // ==========================================
            GroupBox box2 = new GroupBox { Text = "📄 量測項目一覽表 (預覽區)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            _dgvResult = new DataGridView {
                Dock = DockStyle.Top, 
                Height = 600, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, // 🟢 關閉自動填充，改用絕對比例
                RowHeadersVisible = false, 
                Font = new Font("Microsoft JhengHei UI", 11F),
                Margin = new Padding(0, 0, 0, 0), 
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

            box2.Controls.Add(_dgvResult);
            layout.Controls.Add(box2, 0, 1);

            mainScrollPanel.Controls.Add(layout);
            
            BtnSearch_Click(null, null);

            return mainScrollPanel;
        }

        private string GetItemColumnName(DataTable dt)
        {
            List<string> possibleNames = new List<string> { "檢測項目", "項目", "名稱" };
            foreach (string pc in possibleNames) {
                if (dt.Columns.Contains(pc)) return pc;
            }
            return "";
        }

        private string GetPointColumnName(DataTable dt)
        {
            List<string> possibleNames = new List<string> { "檢測點", "檢測名稱", "設備名稱", "點位", "SEG編號", "水錶名稱", "項目", "單位" };
            foreach (string pc in possibleNames) {
                if (dt.Columns.Contains(pc)) return pc;
            }
            return "";
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (_cboYear.SelectedItem == null) return;
            string targetYear = _cboYear.SelectedItem.ToString();
            
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            try 
            {
                _dtResult = new DataTable();
                _dtResult.Columns.Add("量測項目", typeof(string));
                _dtResult.Columns.Add("檢測點", typeof(string));
                _dtResult.Columns.Add("管制標準", typeof(string));
                for (int i = 1; i <= 12; i++) _dtResult.Columns.Add($"{i}月", typeof(string));
                _dtResult.Columns.Add("最大值", typeof(string));
                _dtResult.Columns.Add("最小值", typeof(string));
                _dtResult.Columns.Add("平均值", typeof(string));

                var aggregatedData = new Dictionary<string, Dictionary<int, List<double>>>();

                foreach (string tbName in _targetTables) 
                {
                    DataTable dt = null;
                    try {
                        dt = DataManager.GetTableData("TestData", tbName, "", "", "");
                    } catch { continue; }

                    if (dt == null || dt.Rows.Count == 0) continue;

                    string dateCol = dt.Columns.Contains("日期") ? "日期" : (dt.Columns.Contains("年月") ? "年月" : "");
                    string itemCol = GetItemColumnName(dt);
                    string pointCol = GetPointColumnName(dt);
                    string valCol = dt.Columns.Contains("檢測數據") ? "檢測數據" : "";
                    string limitCol = dt.Columns.Contains("管制值") ? "管制值" : "";

                    if (string.IsNullOrEmpty(dateCol) || string.IsNullOrEmpty(itemCol) || string.IsNullOrEmpty(valCol)) continue;

                    foreach (DataRow r in dt.Rows) 
                    {
                        if (r.RowState == DataRowState.Deleted) continue;

                        string dateStr = r[dateCol]?.ToString() ?? "";
                        if (!dateStr.StartsWith(targetYear)) continue; 

                        int month = 0;
                        if (dateStr.Length >= 7) {
                            string monthStr = dateStr.Substring(5, 2);
                            int.TryParse(monthStr, out month);
                        }

                        if (month < 1 || month > 12) continue;

                        string item = r[itemCol]?.ToString()?.Trim() ?? "";
                        string point = !string.IsNullOrEmpty(pointCol) ? (r[pointCol]?.ToString()?.Trim() ?? "") : "";
                        string limit = !string.IsNullOrEmpty(limitCol) ? (r[limitCol]?.ToString()?.Trim() ?? "") : "";
                        string valStr = r[valCol]?.ToString()?.Replace(",", "").Trim() ?? "";

                        if (string.IsNullOrEmpty(item) || string.IsNullOrEmpty(valStr)) continue;

                        if (double.TryParse(valStr, out double val)) 
                        {
                            string key = $"{item}|{point}|{limit}";

                            if (!aggregatedData.ContainsKey(key)) {
                                aggregatedData[key] = new Dictionary<int, List<double>>();
                                for (int i = 1; i <= 12; i++) aggregatedData[key][i] = new List<double>();
                            }

                            aggregatedData[key][month].Add(val);
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

                    for (int m = 1; m <= 12; m++) 
                    {
                        var mValues = kvp.Value[m];
                        if (mValues.Count > 0) 
                        {
                            row[$"{m}月"] = string.Join("、", mValues.Select(v => v.ToString("0.##")));
                            allValuesForYear.AddRange(mValues);
                        } 
                        else 
                        {
                            row[$"{m}月"] = "";
                        }
                    }

                    if (allValuesForYear.Count > 0) {
                        row["最大值"] = allValuesForYear.Max().ToString("0.##");
                        row["最小值"] = allValuesForYear.Min().ToString("0.##");
                        row["平均值"] = allValuesForYear.Average().ToString("0.##");
                    } else {
                        row["最大值"] = ""; row["最小值"] = ""; row["平均值"] = "";
                    }

                    _dtResult.Rows.Add(row);
                }

                _dgvResult.DataSource = _dtResult;

                // 🟢 強制套用 25 等分比例寬度
                if (_dgvResult.Columns.Count > 0) 
                {
                    int totalGridWidth = _dgvResult.ClientSize.Width;
                    // 如果畫面還沒完全生成導致寬度太小，給定一個合理的基準寬度
                    if (totalGridWidth < 500) totalGridWidth = 1400; 

                    float unitWidth = totalGridWidth / 25f;

                    _dgvResult.Columns["量測項目"].Width = (int)(unitWidth * 5);
                    _dgvResult.Columns["檢測點"].Width = (int)(unitWidth * 3);
                    _dgvResult.Columns["管制標準"].Width = (int)(unitWidth * 2);

                    for (int m = 1; m <= 12; m++) {
                        _dgvResult.Columns[$"{m}月"].Width = (int)unitWidth;
                    }

                    _dgvResult.Columns["最大值"].Width = (int)unitWidth;
                    _dgvResult.Columns["最小值"].Width = (int)unitWidth;
                    _dgvResult.Columns["平均值"].Width = (int)unitWidth;
                }

                _dgvResult.ClearSelection();
            } 
            catch (Exception ex) 
            {
                MessageBox.Show("查詢失敗：" + ex.Message, "錯誤");
            } 
            finally 
            {
                // 🟢 修復：確保即使發生錯誤也會正確解除漏斗等待狀態
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (_dtResult == null || _dtResult.Rows.Count == 0) {
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
                        float[] colWidths = new float[_dtResult.Columns.Count];
                        
                        // 🟢 完美套用 25 等分比例寬度設定
                        float unitWidth = w / 25f;
                        
                        colWidths[0] = unitWidth * 5f; // 量測項目
                        colWidths[1] = unitWidth * 3f; // 檢測點
                        colWidths[2] = unitWidth * 2f; // 管制標準
                        
                        for (int i = 3; i <= 14; i++) colWidths[i] = unitWidth; // 1~12月
                        colWidths[15] = unitWidth; // 最大值
                        colWidths[16] = unitWidth; // 最小值
                        colWidths[17] = unitWidth; // 平均值

                        // 第一層表頭
                        float currX = x;
                        // 左側三大欄位直接畫兩倍高
                        for (int i = 0; i <= 2; i++) {
                            RectangleF rHead = new RectangleF(currX, y, colWidths[i], rowH * 2);
                            g.FillRectangle(Brushes.LightGray, rHead);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rHead));
                            g.DrawString(_dtResult.Columns[i].ColumnName, fGridTitle, Brushes.Black, rHead, sfCenter);
                            currX += colWidths[i];
                        }
                        
                        // 右側群組標題「XX 年度」
                        float monthGroupWidth = unitWidth * 15f; 
                        RectangleF rYearTitle = new RectangleF(currX, y, monthGroupWidth, rowH);
                        g.FillRectangle(Brushes.LightGray, rYearTitle);
                        g.DrawRectangle(Pens.Black, Rectangle.Round(rYearTitle));
                        g.DrawString($"{targetYear} 年度", fGridTitle, Brushes.Black, rYearTitle, sfCenter);

                        // 第二層表頭 (1~12月, 最大, 最小, 平均)
                        float subY = y + rowH;
                        float subX = currX;
                        for (int i = 3; i < _dtResult.Columns.Count; i++) {
                            RectangleF rSubHead = new RectangleF(subX, subY, colWidths[i], rowH);
                            g.FillRectangle(Brushes.LightGray, rSubHead);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rSubHead));
                            g.DrawString(_dtResult.Columns[i].ColumnName, fGridHead, Brushes.Black, rSubHead, sfCenter);
                            subX += colWidths[i];
                        }
                        
                        y += rowH * 2;

                        // 3. 資料清單
                        while (currentRowIndex < _dtResult.Rows.Count) 
                        {
                            DataRow row = _dtResult.Rows[currentRowIndex];
                            
                            if (y + rowH > ev.MarginBounds.Bottom - 30) {
                                g.DrawString("8-ES-B11-02 環境量測項目績效一覽表", fFooter, Brushes.Black, x, ev.MarginBounds.Bottom - 15);
                                g.DrawString($"{pageNumber} / {totalPages}", fFooter, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom - 15, w, 20), sfCenter);
                                pageNumber++;
                                ev.HasMorePages = true;
                                return;
                            }

                            currX = x;
                            for (int i = 0; i < _dtResult.Columns.Count; i++) {
                                RectangleF rCell = new RectangleF(currX, y, colWidths[i], rowH);
                                g.DrawRectangle(Pens.Black, Rectangle.Round(rCell));
                                string val = row[i]?.ToString() ?? "";
                                g.DrawString(val, fGridBody, Brushes.Black, rCell, sfCenter);
                                currX += colWidths[i];
                            }
                            y += rowH;
                            currentRowIndex++;
                        }

                        // 4. 補空白列填滿畫面以符合台玻標準要求
                        while (y + rowH <= ev.MarginBounds.Bottom - 30)
                        {
                            currX = x;
                            for (int i = 0; i < _dtResult.Columns.Count; i++) {
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
                        MessageBox.Show("PDF 導出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 導出失敗：" + ex.Message, "錯誤");
                    }
                }
            }
        }
    }
}
