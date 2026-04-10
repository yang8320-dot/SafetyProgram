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
            try 
            { 
                LoadAndMergeData(); 
            }
            catch (Exception ex) 
            { 
                _errorLogs.Add($"[資料讀取] {ex.Message}"); 
            }

            // 主排版引擎
            TableLayoutPanel mainPanel = new TableLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke, 
                AutoScroll = true, 
                RowCount = 3, 
                ColumnCount = 1,
                Padding = new Padding(20) 
            };

            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 框1: 今年修正法規
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 框2: 統計摘要
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 框3: 目錄清單

            // ==========================================
            // 大框 1：今年修正法規 (移至最上方)
            // ==========================================
            GroupBox box1 = CreateDataBox("📌 今年修正法規一覽 (排除重複名稱，依適用性權重顯示)");
            
            Panel pnlAction1 = CreateActionPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvThisYear, "今年修正法規"), 
                () => ExportToPdf(_dgvThisYear, "今年修正法規", "今年修正法規一覽表"));
            
            _dgvThisYear = CreateStandardGrid();
            try 
            { 
                PopulateThisYearData(_dgvThisYear); 
            } 
            catch (Exception ex) 
            { 
                _errorLogs.Add($"[今年修正] {ex.Message}"); 
            }
            
            box1.Controls.Add(_dgvThisYear);
            box1.Controls.Add(pnlAction1);
            _dgvThisYear.BringToFront(); 

            mainPanel.Controls.Add(box1, 0, 0);

            // ==========================================
            // 大框 2：統計摘要 (包含兩行標題 + 表格)
            // ==========================================
            GroupBox box2 = CreateDataBox("📊 統計摘要");
            
            // 🟢 需求3：修正文字為 法令鑑別統計表
            Label lblTitle2 = new Label 
            { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n法令鑑別統計表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter, 
                ForeColor = Color.DarkSlateBlue,
                Dock = DockStyle.Top, 
                Height = 70 
            };

            Panel pnlAction2 = CreateActionPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvStats, "統計摘要"), 
                () => ExportToPdf(_dgvStats, "統計摘要", "法令鑑別統計表"));
            
            _dgvStats = CreateStatsGrid();
            try 
            { 
                PopulateStatsData(_dgvStats); 
            } 
            catch (Exception ex) 
            { 
                _errorLogs.Add($"[統計摘要] {ex.Message}"); 
            }
            
            box2.Controls.Add(_dgvStats);
            box2.Controls.Add(pnlAction2);
            box2.Controls.Add(lblTitle2);
            _dgvStats.BringToFront();

            mainPanel.Controls.Add(box2, 0, 1);

            // ==========================================
            // 大框 3：目錄清單
            // ==========================================
            GroupBox box3 = CreateDataBox("📋 依類別檢視法令名稱");

            Panel pnlTop3 = new Panel { Dock = DockStyle.Top, Height = 70 };
            Label lblTitle3 = new Label 
            { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n法令及其他要求目錄一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DarkSlateBlue,
                Dock = DockStyle.Fill 
            };
            pnlTop3.Controls.Add(lblTitle3);

            Panel pnlAction3 = CreateActionPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvCategoryLaws, "法令目錄"), 
                () => ExportToPdf(_dgvCategoryLaws, "法令目錄", "法令及其他要求目錄一覽表"));
            
            Label lblCbo = new Label 
            { 
                Text = "選擇類別:", 
                AutoSize = true, 
                Location = new Point(350, 10), 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) 
            };
            
            // 🟢 需求5：調整間距，避免互相遮擋 (X軸由430改為450)
            _cboCategory = new ComboBox 
            { 
                DropDownStyle = ComboBoxStyle.DropDownList, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                Width = 180, 
                Location = new Point(450, 6) 
            };
            
            _cboCategory.Items.AddRange(_tableNames);
            _cboCategory.SelectedIndexChanged += (s, e) => 
            { 
                try { PopulateCategoryLaws(); } 
                catch { } 
            };
            
            pnlAction3.Controls.Add(lblCbo);
            pnlAction3.Controls.Add(_cboCategory);

            _dgvCategoryLaws = CreateStandardGrid();
            
            box3.Controls.Add(_dgvCategoryLaws);
            box3.Controls.Add(pnlAction3);
            box3.Controls.Add(pnlTop3);
            _dgvCategoryLaws.BringToFront();

            mainPanel.Controls.Add(box3, 0, 2);

            if (_cboCategory.Items.Count > 0) 
            {
                _cboCategory.SelectedIndex = 0;
            }

            if (_errorLogs.Count > 0) 
            {
                MessageBox.Show("部分資料讀取異常，系統已略過。\n" + string.Join("\n", _errorLogs), "通知", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return mainPanel;
        }

        // ==========================================
        // UI 元件工廠方法
        // ==========================================
        private GroupBox CreateDataBox(string title)
        {
            return new GroupBox 
            { 
                Text = title, 
                Dock = DockStyle.Fill, 
                MinimumSize = new Size(0, 350), 
                Margin = new Padding(0, 0, 0, 20), 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Padding = new Padding(15) 
            };
        }

        private Panel CreateActionPanel(string exText, string pdfText, Action exClick, Action pdfClick)
        {
            Panel p = new Panel { Dock = DockStyle.Top, Height = 45 };
            
            Button btnEx = new Button 
            { 
                Text = "📊 " + exText, 
                Size = new Size(150, 32), 
                Location = new Point(10, 5), 
                BackColor = Color.MediumSeaGreen, 
                ForeColor = Color.White, 
                Cursor = Cursors.Hand 
            };
            
            Button btnPdf = new Button 
            { 
                Text = "📄 " + pdfText, 
                Size = new Size(150, 32), 
                Location = new Point(170, 5), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Cursor = Cursors.Hand 
            };
            
            btnEx.Click += (s, e) => exClick();
            btnPdf.Click += (s, e) => pdfClick();
            
            p.Controls.Add(btnEx);
            p.Controls.Add(btnPdf);
            return p;
        }

        private DataGridView CreateStandardGrid()
        {
            return new DataGridView 
            { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false, 
                ReadOnly = true, 
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                Font = new Font("Microsoft JhengHei UI", 11F), 
                BorderStyle = BorderStyle.Fixed3D, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, 
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } 
            };
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
            foreach (var col in expectedCols) 
            {
                _dtAllLaws.Columns.Add(col, typeof(string));
            }

            foreach (string tbl in _tableNames) 
            {
                try 
                {
                    DataTable dt = DataManager.GetTableData(DbName, tbl, "", "", "");
                    if (dt == null || dt.Rows.Count == 0) continue; 
                    
                    foreach (DataRow row in dt.Rows) 
                    {
                        DataRow newRow = _dtAllLaws.NewRow();
                        newRow["主分類"] = tbl;
                        foreach (string col in expectedCols) 
                        {
                            if (dt.Columns.Contains(col) && row[col] != DBNull.Value)
                            {
                                newRow[col] = row[col].ToString();
                            }
                            else
                            {
                                newRow[col] = "";
                            }
                        }
                        _dtAllLaws.Rows.Add(newRow);
                    }
                } 
                catch { }
            }
        }

        private string GetSafeStr(DataRowView row, string colName) 
        { 
            if (row.Row.Table.Columns.Contains(colName) && row[colName] != DBNull.Value)
            {
                return row[colName].ToString().Trim();
            }
            return ""; 
        }

        private void PopulateThisYearData(DataGridView dgv)
        {
            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("日期"); 
            dtShow.Columns.Add("鑑別日期"); 
            dtShow.Columns.Add("法規名稱"); 
            dtShow.Columns.Add("適用性");

            if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0) 
            {
                string currentYear = DateTime.Now.Year.ToString();
                DataView dv = new DataView(_dtAllLaws); 
                dv.RowFilter = $"日期 LIKE '%{currentYear}%'";

                Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
                
                foreach (DataRowView drv in dv) 
                {
                    string lawName = GetSafeStr(drv, "法規名稱");
                    if (string.IsNullOrEmpty(lawName)) continue;
                    
                    if (!groupedData.ContainsKey(lawName)) 
                    {
                        groupedData[lawName] = new List<DataRowView>();
                    }
                    groupedData[lawName].Add(drv);
                }

                foreach (var kvp in groupedData) 
                {
                    string latestDate = "";
                    string latestIdenDate = "";
                    string firstApply = GetSafeStr(kvp.Value[0], "適用性");
                    bool hasApplicable = false;
                    
                    foreach (var row in kvp.Value) 
                    {
                        string d = GetSafeStr(row, "日期");
                        string iden = GetSafeStr(row, "鑑別日期");
                        string apply = GetSafeStr(row, "適用性");
                        
                        if (string.Compare(d, latestDate) > 0) latestDate = d;
                        if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                        if (apply == "適用") hasApplicable = true;
                    }
                    
                    dtShow.Rows.Add(latestDate, latestIdenDate, kvp.Key, hasApplicable ? "適用" : firstApply);
                }
            }
            
            dgv.DataSource = dtShow;
            if (dgv.Columns.Count >= 4) 
            {
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void PopulateStatsData(DataGridView dgv)
        {
            DataTable dtStats = new DataTable();
            string[] cols = { "類別", "收集【法規】數", "收集【法條】數", "【適用】法條數", "【參考】數", "【不適用】法條數", "【確認中】數", "合法且有提升\n績效機會法條數", "合法但潛在不\n符合風險法條數", "【未鑑別】\n法條數" };
            foreach (string c in cols) 
            {
                dtStats.Columns.Add(c);
            }

            int[] sums = new int[9];
            foreach (string cat in _tableNames) 
            {
                int[] v = new int[9];
                if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0) 
                {
                    DataView dv = new DataView(_dtAllLaws); 
                    dv.RowFilter = $"主分類 = '{cat}'";
                    
                    v[1] = dv.Count; 
                    HashSet<string> uniqueNames = new HashSet<string>();
                    
                    foreach (DataRowView row in dv) 
                    {
                        string name = GetSafeStr(row, "法規名稱");
                        string aStatus = GetSafeStr(row, "適用性");
                        string cStatus = GetSafeStr(row, "合規狀態");
                        
                        if (!string.IsNullOrEmpty(name)) uniqueNames.Add(name);
                        
                        if (aStatus == "適用") v[2]++; 
                        else if (aStatus == "參考") v[3]++; 
                        else if (aStatus == "不適用") v[4]++; 
                        else if (aStatus == "確認中") v[5]++; 
                        else if (string.IsNullOrEmpty(aStatus)) v[8]++;
                        
                        if (cStatus.Contains("提升")) v[6]++; 
                        if (cStatus.Contains("潛在不符合")) v[7]++;
                    }
                    v[0] = uniqueNames.Count; 
                }
                dtStats.Rows.Add(cat, v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8]);
                for (int i = 0; i < 9; i++) 
                {
                    sums[i] += v[i];
                }
            }
            
            dtStats.Rows.Add("合計", sums[0], sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], sums[8]);
            dgv.DataSource = dtStats;
            
            if (dgv.Columns.Count > 0) 
            {
                dgv.Columns[0].HeaderText = "環保法規\n(類別)";
            }
            
            dgv.DataBindingComplete += (s, e) => { 
                if (dgv.Rows.Count > 0) 
                {
                    dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold); 
                }
            };
        }

        // 🟢 需求4：直接從「法規目錄一覽」抓取資料，過濾選項類別
        private void PopulateCategoryLaws()
        {
            if (_cboCategory.SelectedItem == null) return;
            string category = _cboCategory.SelectedItem.ToString();
            
            DataTable dtShow = new DataTable();
            try 
            {
                DataTable dtDir = DataManager.GetTableData(DbName, "法規目錄一覽", "", "", "");
                if (dtDir != null && dtDir.Columns.Count > 0)
                {
                    dtShow = dtDir.Clone(); 
                    foreach (DataRow row in dtDir.Rows) 
                    {
                        if (row["選項類別"]?.ToString() == category) 
                        {
                            dtShow.ImportRow(row);
                        }
                    }
                }
            }
            catch 
            {
                // 若法規目錄一覽尚未建立會觸發 Exception，給予空表避免當機
            }

            _dgvCategoryLaws.DataSource = dtShow;
            
            // 隱藏 Id, 選項類別
            if (_dgvCategoryLaws.Columns.Contains("Id")) _dgvCategoryLaws.Columns["Id"].Visible = false;
            if (_dgvCategoryLaws.Columns.Contains("選項類別")) _dgvCategoryLaws.Columns["選項類別"].Visible = false;

            // 調整顯示欄位寬度
            if (_dgvCategoryLaws.Columns.Contains("流水號")) _dgvCategoryLaws.Columns["流水號"].Width = 80;
            if (_dgvCategoryLaws.Columns.Contains("法規名稱")) _dgvCategoryLaws.Columns["法規名稱"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            if (_dgvCategoryLaws.Columns.Contains("日期")) _dgvCategoryLaws.Columns["日期"].Width = 120;
            if (_dgvCategoryLaws.Columns.Contains("適用性")) _dgvCategoryLaws.Columns["適用性"].Width = 100;
            if (_dgvCategoryLaws.Columns.Contains("鑑別日期")) _dgvCategoryLaws.Columns["鑑別日期"].Width = 120;
            if (_dgvCategoryLaws.Columns.Contains("再次確認日期")) _dgvCategoryLaws.Columns["再次確認日期"].Width = 150;
        }

        // ==========================================
        // 匯出功能 (Excel / PDF)
        // ==========================================
        private void ExportToExcel(DataGridView dgv, string title)
        {
            if (dgv.Rows.Count == 0) 
            { 
                MessageBox.Show("沒有資料可匯出！"); 
                return; 
            }
            
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        List<DataGridViewColumn> visCols = new List<DataGridViewColumn>();

                        // 只匯出可見欄位 (避開 Id, 選項類別)
                        foreach (DataGridViewColumn col in dgv.Columns) {
                            if (col.Visible) {
                                visCols.Add(col);
                                dt.Columns.Add(col.HeaderText);
                            }
                        }
                        
                        foreach (DataGridViewRow row in dgv.Rows) {
                            if (row.IsNewRow) continue;
                            DataRow dRow = dt.NewRow();
                            for (int i = 0; i < visCols.Count; i++) {
                                var cellVal = row.Cells[visCols[i].Index].Value;
                                dRow[i] = cellVal != null ? cellVal.ToString() : "";
                            }
                            dt.Rows.Add(dRow);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("Data");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns(); 
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("Excel 匯出成功！");
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("Excel 匯出失敗：" + ex.Message); 
                    }
                }
            }
        }

        // 🟢 需求 1 & 2：匯出 PDF (新增 reportTitle 參數以支援表頭與導出日期標示)
        private void ExportToPdf(DataGridView dgv, string fileName, string reportTitle)
        {
            if (dgv.Rows.Count == 0) 
            { 
                MessageBox.Show("沒有資料可列印！"); 
                return; 
            }
            
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = fileName + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true; 
                    pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                    int rowIndex = 0;
                    pd.PrintPage += (s, e) => 
                    {
                        Graphics g = e.Graphics;
                        Font font = new Font("Microsoft JhengHei UI", 9F);
                        Font headerFont = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
                        Font titleFont = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                        Font dateFont = new Font("Microsoft JhengHei UI", 11F);
                        
                        StringFormat fmtCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                        StringFormat fmtLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };

                        float y = e.MarginBounds.Top;
                        float totalWidth = 0;
                        
                        // 計算可見欄位寬度
                        List<DataGridViewColumn> visCols = new List<DataGridViewColumn>();
                        foreach (DataGridViewColumn col in dgv.Columns) {
                            if (col.Visible) {
                                visCols.Add(col);
                                totalWidth += col.Width;
                            }
                        }

                        float scale = e.MarginBounds.Width / totalWidth;
                        if (scale > 1f) scale = 1f; 

                        g.ScaleTransform(scale, scale);
                        float scaledHeight = e.MarginBounds.Height / scale;
                        float scaledWidth = e.MarginBounds.Width / scale;
                        float x = e.MarginBounds.Left / scale;

                        // === 繪製 PDF 表頭 (公司與標題、導出日期) ===
                        string companyTitle = "台灣玻璃工業股份有限公司-彰濱廠\n" + reportTitle;
                        string exportDate = "導出日期: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                        // 畫大標題
                        SizeF titleSize = g.MeasureString(companyTitle, titleFont, (int)scaledWidth, fmtCenter);
                        RectangleF titleRect = new RectangleF(x, y / scale, scaledWidth, titleSize.Height + 10);
                        g.DrawString(companyTitle, titleFont, Brushes.DarkSlateBlue, titleRect, fmtCenter);
                        y += (titleSize.Height + 10) * scale;

                        // 畫副標題 (導出日期)
                        SizeF dateSize = g.MeasureString(exportDate, dateFont, (int)scaledWidth, fmtLeft);
                        RectangleF dateRect = new RectangleF(x, y / scale, scaledWidth, dateSize.Height + 10);
                        g.DrawString(exportDate, dateFont, Brushes.Black, dateRect, fmtLeft);
                        y += (dateSize.Height + 15) * scale;
                        // ===========================================

                        float headerH = dgv.ColumnHeadersHeight < 40 ? 40 : dgv.ColumnHeadersHeight;
                        
                        // 繪製表格 Header
                        for (int i = 0; i < visCols.Count; i++) 
                        {
                            RectangleF rectF = new RectangleF(x, y / scale, visCols[i].Width, headerH);
                            Rectangle rect = Rectangle.Round(rectF); 
                            
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect);
                            
                            string headerText = visCols[i].HeaderText.Replace("\n", "");
                            g.DrawString(headerText, headerFont, Brushes.Black, rect, fmtCenter);
                            x += visCols[i].Width;
                        }
                        y += headerH * scale;

                        // 繪製資料
                        while (rowIndex < dgv.Rows.Count) 
                        {
                            DataGridViewRow row = dgv.Rows[rowIndex];
                            float rowH = row.Height < 30 ? 30 : row.Height;
                            
                            // 若超出版面高度，換頁
                            if ((y / scale) + rowH > scaledHeight + (e.MarginBounds.Top / scale)) 
                            {
                                e.HasMorePages = true; 
                                return;
                            }

                            x = e.MarginBounds.Left / scale;
                            for (int i = 0; i < visCols.Count; i++) 
                            {
                                RectangleF rectF = new RectangleF(x, y / scale, visCols[i].Width, rowH);
                                Rectangle rect = Rectangle.Round(rectF);
                                
                                g.DrawRectangle(Pens.Black, rect);
                                
                                string val = row.Cells[visCols[i].Index].Value?.ToString() ?? "";
                                g.DrawString(val, font, Brushes.Black, rect, fmtCenter);
                                x += visCols[i].Width;
                            }
                            y += rowH * scale;
                            rowIndex++;
                        }
                        e.HasMorePages = false;
                        rowIndex = 0; 
                    };

                    try 
                    { 
                        pd.Print(); 
                        MessageBox.Show("PDF 匯出成功！"); 
                    }
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message); 
                    }
                }
            }
        }
    }
}
