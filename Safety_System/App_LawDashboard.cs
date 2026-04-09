/// FILE: Safety_System/App_LawDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace Safety_System
{
    public class App_LawDashboard
    {
        private const string DbName = "法規";
        private readonly string[] _tableNames = { "環保法規", "職安衛法規", "其它法規" };
        
        // 快取資料表
        private DataTable _dtAllLaws;
        private List<string> _errorLogs = new List<string>();
        
        // 資料框 UI 控制項
        private ComboBox _cboCategory;
        private DataGridView _dgvStats;
        private DataGridView _dgvCategoryLaws;
        private DataGridView _dgvThisYear;

        public Control GetView()
        {
            _errorLogs.Clear();
            try { LoadAndMergeData(); }
            catch (Exception ex) { _errorLogs.Add($"[資料讀取] {ex.Message}"); }

            // 🟢 主排版引擎：強制劃分 5 個橫列 (Row)，確保版面絕對不跑位
            TableLayoutPanel mainPanel = new TableLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke, 
                AutoScroll = true, 
                RowCount = 5, 
                ColumnCount = 1,
                Padding = new Padding(20) 
            };

            // 設定 5 個框的高度屬性 (文字框固定高度，資料框自動延展)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));  // 框1: 統計標題
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 框2: 統計表格
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));  // 框3: 目錄標題
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 框4: 目錄表格
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 框5: 今年法規表格

            // ==========================================
            // 框 1 (純文字)：統計摘要標題 (原紅框標題)
            // ==========================================
            GroupBox boxText1 = CreateTextBox("台灣玻璃工業股份有限公司-彰濱廠\n環安衛法令及其他要求內容一覽表");
            mainPanel.Controls.Add(boxText1, 0, 0);

            // ==========================================
            // 框 2 (資料庫)：統計摘要表格
            // ==========================================
            GroupBox boxData1 = CreateDataBox("📊 統計摘要");
            Panel pnlAction1 = CreateActionPanel("匯出統計 Excel", "匯出統計 PDF", () => ExportToExcel(_dgvStats, "統計摘要"), () => ExportToPdf(_dgvStats, "統計摘要"));
            _dgvStats = CreateStatsGrid();
            try { PopulateStatsData(_dgvStats); } catch (Exception ex) { _errorLogs.Add($"[統計摘要] {ex.Message}"); }
            boxData1.Controls.Add(_dgvStats);
            boxData1.Controls.Add(pnlAction1);
            mainPanel.Controls.Add(boxData1, 0, 1);

            // ==========================================
            // 框 3 (純文字)：目錄清單標題 (原藍框標題)
            // ==========================================
            GroupBox boxText2 = CreateTextBox("台灣玻璃工業股份有限公司-彰濱廠\n環安衛法令及其他要求目錄一覽表");
            mainPanel.Controls.Add(boxText2, 0, 2);

            // ==========================================
            // 框 4 (資料庫)：依類別檢視目錄
            // ==========================================
            GroupBox boxData2 = CreateDataBox("📋 依類別檢視法令名稱");
            Panel pnlAction2 = CreateActionPanel("匯出目錄 Excel", "匯出目錄 PDF", () => ExportToExcel(_dgvCategoryLaws, "法令目錄"), () => ExportToPdf(_dgvCategoryLaws, "法令目錄"));
            
            // 下拉選單加入工具列
            _cboCategory = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 200, Location = new Point(10, 8) };
            _cboCategory.Items.AddRange(_tableNames);
            _cboCategory.SelectedIndexChanged += (s, e) => { try { PopulateCategoryLaws(); } catch { } };
            pnlAction2.Controls.Add(new Label { Text = "選擇類別:", AutoSize = true, Location = new Point(220, 12), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) });
            pnlAction2.Controls.Add(_cboCategory);
            _cboCategory.Left = 300;

            _dgvCategoryLaws = CreateStandardGrid();
            boxData2.Controls.Add(_dgvCategoryLaws);
            boxData2.Controls.Add(pnlAction2);
            mainPanel.Controls.Add(boxData2, 0, 3);

            // ==========================================
            // 框 5 (資料庫)：今年修正法規
            // ==========================================
            GroupBox boxData3 = CreateDataBox("📌 今年修正法規一覽 (排除重複名稱，依適用性權重顯示)");
            Panel pnlAction3 = CreateActionPanel("匯出法規 Excel", "匯出法規 PDF", () => ExportToExcel(_dgvThisYear, "今年修正法規"), () => ExportToPdf(_dgvThisYear, "今年修正法規"));
            _dgvThisYear = CreateStandardGrid();
            try { PopulateThisYearData(_dgvThisYear); } catch (Exception ex) { _errorLogs.Add($"[今年修正] {ex.Message}"); }
            boxData3.Controls.Add(_dgvThisYear);
            boxData3.Controls.Add(pnlAction3);
            mainPanel.Controls.Add(boxData3, 0, 4);

            // 初始化選項
            if (_cboCategory.Items.Count > 0) _cboCategory.SelectedIndex = 0;

            // 錯誤提示
            if (_errorLogs.Count > 0) MessageBox.Show("部分資料讀取異常，系統已略過。\n" + string.Join("\n", _errorLogs), "通知", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            return mainPanel;
        }

        // ==========================================
        // UI 元件工廠方法
        // ==========================================
        private GroupBox CreateTextBox(string text)
        {
            GroupBox gb = new GroupBox { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(5) };
            Label lbl = new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };
            gb.Controls.Add(lbl);
            return gb;
        }

        private GroupBox CreateDataBox(string title)
        {
            return new GroupBox { Text = title, Dock = DockStyle.Fill, MinimumSize = new Size(0, 350), Margin = new Padding(0, 0, 0, 20), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
        }

        private Panel CreateActionPanel(string exText, string pdfText, Action exClick, Action pdfClick)
        {
            Panel p = new Panel { Dock = DockStyle.Top, Height = 45 };
            Button btnEx = new Button { Text = "📊 " + exText, Size = new Size(160, 32), Location = new Point(10, 5), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand };
            Button btnPdf = new Button { Text = "📄 " + pdfText, Size = new Size(160, 32), Location = new Point(180, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand };
            
            btnEx.Click += (s, e) => exClick();
            btnPdf.Click += (s, e) => pdfClick();
            
            p.Controls.Add(btnEx);
            p.Controls.Add(btnPdf);
            btnEx.BringToFront(); btnPdf.BringToFront();
            return p;
        }

        private DataGridView CreateStandardGrid()
        {
            return new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, AllowUserToDeleteRows = false, ReadOnly = true, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, Font = new Font("Microsoft JhengHei UI", 11F), BorderStyle = BorderStyle.Fixed3D, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } };
        }

        private DataGridView CreateStatsGrid()
        {
            DataGridView dgv = CreateStandardGrid();
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.YellowGreen;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
            dgv.DefaultCellStyle.SelectionBackColor = Color.Khaki;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.GridColor = Color.Black;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            return dgv;
        }

        // ==========================================
        // 資料載入與運算邏輯
        // ==========================================
        private void LoadAndMergeData()
        {
            _dtAllLaws = new DataTable();
            _dtAllLaws.Columns.Add("主分類", typeof(string));
            string[] expectedCols = { "Id", "日期", "法規名稱", "發布機關", "施行日期", "合規狀態", "適用性", "鑑別日期" };
            foreach (var col in expectedCols) _dtAllLaws.Columns.Add(col, typeof(string));

            foreach (string tbl in _tableNames) {
                try {
                    DataTable dt = DataManager.GetTableData(DbName, tbl, "", "", "");
                    if (dt == null || dt.Rows.Count == 0) continue; 
                    foreach (DataRow row in dt.Rows) {
                        DataRow newRow = _dtAllLaws.NewRow();
                        newRow["主分類"] = tbl;
                        foreach (string col in expectedCols) newRow[col] = (dt.Columns.Contains(col) && row[col] != DBNull.Value) ? row[col].ToString() : "";
                        _dtAllLaws.Rows.Add(newRow);
                    }
                } catch { }
            }
        }

        private string GetSafeStr(DataRowView row, string colName) { return (row.Row.Table.Columns.Contains(colName) && row[colName] != DBNull.Value) ? row[colName].ToString().Trim() : ""; }

        private void PopulateThisYearData(DataGridView dgv)
        {
            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("日期"); dtShow.Columns.Add("鑑別日期"); dtShow.Columns.Add("法規名稱"); dtShow.Columns.Add("適用性");

            if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0) {
                string currentYear = DateTime.Now.Year.ToString();
                DataView dv = new DataView(_dtAllLaws); dv.RowFilter = $"日期 LIKE '%{currentYear}%'";

                Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
                foreach (DataRowView drv in dv) {
                    string lawName = GetSafeStr(drv, "法規名稱");
                    if (string.IsNullOrEmpty(lawName)) continue;
                    if (!groupedData.ContainsKey(lawName)) groupedData[lawName] = new List<DataRowView>();
                    groupedData[lawName].Add(drv);
                }

                foreach (var kvp in groupedData) {
                    string latestDate = "", latestIdenDate = "", firstApply = GetSafeStr(kvp.Value[0], "適用性");
                    bool hasApplicable = false;
                    foreach (var row in kvp.Value) {
                        string d = GetSafeStr(row, "日期"), iden = GetSafeStr(row, "鑑別日期"), apply = GetSafeStr(row, "適用性");
                        if (string.Compare(d, latestDate) > 0) latestDate = d;
                        if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                        if (apply == "適用") hasApplicable = true;
                    }
                    dtShow.Rows.Add(latestDate, latestIdenDate, kvp.Key, hasApplicable ? "適用" : firstApply);
                }
            }
            dgv.DataSource = dtShow;
            if (dgv.Columns.Count >= 4) dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void PopulateStatsData(DataGridView dgv)
        {
            DataTable dtStats = new DataTable();
            string[] cols = { "類別", "收集【法規】數", "收集【法條】數", "【適用】法條數", "【參考】數", "【不適用】法條數", "【確認中】數", "合法且有提升\n績效機會法條數", "合法但潛在不\n符合風險法條數", "【未鑑別】\n法條數" };
            foreach (string c in cols) dtStats.Columns.Add(c);

            int[] sums = new int[9];
            foreach (string cat in _tableNames) {
                int[] v = new int[9];
                if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0) {
                    DataView dv = new DataView(_dtAllLaws); dv.RowFilter = $"主分類 = '{cat}'";
                    v[1] = dv.Count; // 法條數
                    HashSet<string> uniqueNames = new HashSet<string>();
                    foreach (DataRowView row in dv) {
                        string name = GetSafeStr(row, "法規名稱"), aStatus = GetSafeStr(row, "適用性"), cStatus = GetSafeStr(row, "合規狀態");
                        if (!string.IsNullOrEmpty(name)) uniqueNames.Add(name);
                        if (aStatus == "適用") v[2]++; else if (aStatus == "參考") v[3]++; else if (aStatus == "不適用") v[4]++; else if (aStatus == "確認中") v[5]++; else if (string.IsNullOrEmpty(aStatus)) v[8]++;
                        if (cStatus.Contains("提升")) v[6]++; if (cStatus.Contains("潛在不符合")) v[7]++;
                    }
                    v[0] = uniqueNames.Count; // 法規數
                }
                dtStats.Rows.Add(cat, v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8]);
                for (int i = 0; i < 9; i++) sums[i] += v[i];
            }
            dtStats.Rows.Add("合計", sums[0], sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], sums[8]);
            dgv.DataSource = dtStats;
            if (dgv.Columns.Count > 0) dgv.Columns[0].HeaderText = "環保法規\n(類別)";
            dgv.DataBindingComplete += (s, e) => { if (dgv.Rows.Count > 0) dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold); };
        }

        private void PopulateCategoryLaws()
        {
            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("流水號"); dtShow.Columns.Add("法令名稱"); dtShow.Columns.Add("公告日"); dtShow.Columns.Add("鑑別日期");

            if (_cboCategory.SelectedItem != null && _dtAllLaws != null && _dtAllLaws.Rows.Count > 0) {
                DataView dv = new DataView(_dtAllLaws); dv.RowFilter = $"主分類 = '{_cboCategory.SelectedItem}'";
                Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
                foreach (DataRowView drv in dv) {
                    string lawName = GetSafeStr(drv, "法規名稱");
                    if (string.IsNullOrEmpty(lawName)) continue;
                    if (!groupedData.ContainsKey(lawName)) groupedData[lawName] = new List<DataRowView>();
                    groupedData[lawName].Add(drv);
                }

                int index = 1;
                foreach (var kvp in groupedData) {
                    string latestDate = "", latestIdenDate = "";
                    foreach (var row in kvp.Value) {
                        string d = GetSafeStr(row, "日期"), iden = GetSafeStr(row, "鑑別日期");
                        if (string.Compare(d, latestDate) > 0) latestDate = d;
                        if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                    }
                    dtShow.Rows.Add(index++, kvp.Key, latestDate, latestIdenDate);
                }
            }
            _dgvCategoryLaws.DataSource = dtShow;
            if (_dgvCategoryLaws.Columns.Count >= 4) { _dgvCategoryLaws.Columns[0].Width = 80; _dgvCategoryLaws.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; _dgvCategoryLaws.Columns[2].Width = 150; _dgvCategoryLaws.Columns[3].Width = 150; }
        }

        // ==========================================
        // 匯出功能 (Excel 與 PDF)
        // ==========================================
        private void ExportToExcel(DataGridView dgv, string title)
        {
            if (dgv.Rows.Count == 0) { MessageBox.Show("沒有資料可匯出！"); return; }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("Data");
                            for (int i = 0; i < dgv.Columns.Count; i++) ws.Cells[1, i + 1].Value = dgv.Columns[i].HeaderText;
                            for (int i = 0; i < dgv.Rows.Count; i++)
                                for (int j = 0; j < dgv.Columns.Count; j++) ws.Cells[i + 2, j + 1].Value = dgv.Rows[i].Cells[j].Value?.ToString();
                            ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("Excel 匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("Excel 匯出失敗：" + ex.Message); }
                }
            }
        }

        // 🟢 原生 A4 PDF 繪圖匯出引擎
        private void ExportToPdf(DataGridView dgv, string title)
        {
            if (dgv.Rows.Count == 0) { MessageBox.Show("沒有資料可列印！"); return; }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true; // 橫印較適合表格
                    pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                    int rowIndex = 0;
                    pd.PrintPage += (s, e) => {
                        Graphics g = e.Graphics;
                        Font font = new Font("Microsoft JhengHei UI", 9F);
                        Font headerFont = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
                        StringFormat fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };

                        float y = e.MarginBounds.Top;
                        float[] widths = new float[dgv.Columns.Count];
                        float totalWidth = 0;
                        for (int i = 0; i < dgv.Columns.Count; i++) { widths[i] = dgv.Columns[i].Width; totalWidth += widths[i]; }
                        
                        // 依照 A4 寬度等比例縮放表格
                        float scale = e.MarginBounds.Width / totalWidth;
                        if (scale > 1f) scale = 1f; // 若表太小不放大，只縮小

                        g.ScaleTransform(scale, scale);
                        float scaledHeight = e.MarginBounds.Height / scale;

                        // 畫標題列
                        float x = e.MarginBounds.Left / scale;
                        float headerH = dgv.ColumnHeadersHeight < 40 ? 40 : dgv.ColumnHeadersHeight;
                        for (int i = 0; i < dgv.Columns.Count; i++) {
                            RectangleF rect = new RectangleF(x, y / scale, widths[i], headerH);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            g.DrawString(dgv.Columns[i].HeaderText.Replace("\n", ""), headerFont, Brushes.Black, rect, fmt);
                            x += widths[i];
                        }
                        y += headerH * scale;

                        // 畫資料列
                        while (rowIndex < dgv.Rows.Count) {
                            DataGridViewRow row = dgv.Rows[rowIndex];
                            float rowH = row.Height < 30 ? 30 : row.Height;
                            
                            // 若超過頁面高度，換頁
                            if ((y / scale) + rowH > scaledHeight + (e.MarginBounds.Top / scale)) {
                                e.HasMorePages = true; return;
                            }

                            x = e.MarginBounds.Left / scale;
                            for (int i = 0; i < dgv.Columns.Count; i++) {
                                RectangleF rect = new RectangleF(x, y / scale, widths[i], rowH);
                                g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                                string val = row.Cells[i].Value?.ToString() ?? "";
                                g.DrawString(val, font, Brushes.Black, rect, fmt);
                                x += widths[i];
                            }
                            y += rowH * scale;
                            rowIndex++;
                        }
                        e.HasMorePages = false;
                        rowIndex = 0; // 重置
                    };

                    try { pd.Print(); MessageBox.Show("PDF 匯出成功！\n(已儲存至您選擇的路徑)"); }
                    catch (Exception ex) { MessageBox.Show("PDF 匯出失敗 (請確認您的電腦有啟用 'Microsoft Print to PDF')：\n" + ex.Message); }
                }
            }
        }
    }
}
