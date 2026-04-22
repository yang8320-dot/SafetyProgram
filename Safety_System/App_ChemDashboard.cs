/// FILE: Safety_System/App_ChemDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemDashboard
    {
        // UI 控制項成員
        private DataGridView _dgvSDS;
        
        // 常數定義
        private const string DbName = "Chemical";
        private const string TableName = "SDS_Inventory";
        
        // 🟢 更改檔名為 v2，確保舊的快取設定不會覆蓋新的預設 7 欄位版面
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChemSDS_Visibility_v2.txt");
        
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        // 🟢 更新為您指定的 7 個預設顯示欄位與排序
        private readonly string[] _defaultVisibleCols = { "項次", "廠內編號", "化學物質名稱", "危害標示", "供應商", "供應商電話", "SDS版本日期" };

        public Control GetView()
        {
            LoadVisibilitySettings();

            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(20), 
                RowCount = 2,
                BackColor = Color.WhiteSmoke
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 15) 
            };
            
            Button btnPdf = new Button { 
                Text = "📤 導出 SDS 清冊 PDF", 
                Size = new Size(220, 45), 
                BackColor = Color.DarkCyan, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand 
            };
            btnPdf.Click += (s, e) => ExportToPdf();

            Button btnHazardousPdf = new Button { 
                Text = "📄 導出：危害性化學品清單", 
                Size = new Size(260, 45), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                Margin = new Padding(15, 0, 0, 0)
            };
            btnHazardousPdf.Click += (s, e) => ExportToHazardousListPdfDirectly();

            Button btnSettings = new Button { 
                Text = "⚙️ 設定顯示欄位", 
                Size = new Size(180, 45), 
                BackColor = Color.LightSlateGray, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand, 
                Margin = new Padding(15, 0, 0, 0) 
            };
            btnSettings.Click += (s, e) => OpenColumnSettings();

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(btnHazardousPdf);
            pnlAction.Controls.Add(btnSettings);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            GroupBox boxMain = new GroupBox { 
                Text = "📋 化學品管理綜合看板", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                Padding = new Padding(15) 
            };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F)); 
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  

            Panel pnlTitle = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.FromArgb(240, 245, 250), 
                Margin = new Padding(0, 0, 0, 15) 
            };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.LightSkyBlue, ButtonBorderStyle.Solid);
            
            Label lblCompany = new Label { 
                Text = "台灣玻璃工業股份有限公司 - 彰濱廠", 
                Dock = DockStyle.Top, 
                Height = 45, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), 
                ForeColor = Color.SteelBlue 
            };
            
            Label lblSubTitle = new Label { 
                Text = "化學品清單一覽表", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.FromArgb(60, 60, 60) 
            };
            pnlTitle.Controls.Add(lblSubTitle);
            pnlTitle.Controls.Add(lblCompany);

            GroupBox boxGrid = new GroupBox { 
                Text = "SDS 安全資料表庫存明細", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) 
            };
            _dgvSDS = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = true, 
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                BorderStyle = BorderStyle.None 
            };
            _dgvSDS.RowTemplate.Height = 35;
            _dgvSDS.EnableHeadersVisualStyles = false;
            _dgvSDS.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
            _dgvSDS.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvSDS.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            _dgvSDS.ColumnHeadersHeight = 40;
            _dgvSDS.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            boxGrid.Controls.Add(_dgvSDS);

            innerTable.Controls.Add(pnlTitle, 0, 0);
            innerTable.Controls.Add(boxGrid, 0, 1);
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            LoadData();

            return mainLayout;
        }

        private void LoadData()
        {
            try
            {
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                if (dt != null)
                {
                    // 🟢 動態加入「項次」欄位，並填入流水號
                    if (!dt.Columns.Contains("項次"))
                    {
                        DataColumn seqCol = new DataColumn("項次", typeof(int));
                        dt.Columns.Add(seqCol);
                        seqCol.SetOrdinal(0); 

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dt.Rows[i]["項次"] = i + 1;
                        }
                    }

                    _dgvSDS.DataSource = dt;
                    
                    if (_dgvSDS.Columns.Contains("項次")) {
                        _dgvSDS.Columns["項次"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        _dgvSDS.Columns["項次"].Width = 60;
                        _dgvSDS.Columns["項次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    ApplyVisibility(); 
                }
                else
                {
                    _dgvSDS.DataSource = null;
                }
            }
            catch { _dgvSDS.DataSource = null; }
        }

        private void ApplyVisibility()
        {
            if (_dgvSDS.DataSource == null || _dgvSDS.Columns.Count == 0) return;
            
            bool isFirstTime = _columnVisibility.Count == 0;

            foreach (DataGridViewColumn col in _dgvSDS.Columns)
            {
                // 強制隱藏 Id 欄位
                if (col.Name.Equals("Id", StringComparison.OrdinalIgnoreCase)) 
                { 
                    col.Visible = false; 
                    continue; 
                }
                
                // 首次設定時，僅顯示預設欄位；後續則依據使用者的設定
                if (isFirstTime)
                {
                    col.Visible = _defaultVisibleCols.Contains(col.Name);
                    _columnVisibility[col.Name] = col.Visible;
                }
                else
                {
                    col.Visible = _columnVisibility.ContainsKey(col.Name) ? _columnVisibility[col.Name] : false;
                }
            }

            // 🟢 強制處理欄位排序
            int currentIndex = 0;
            
            // 1. 先排序「預設欄位清單」中的欄位 (若其狀態為顯示)
            foreach (string defCol in _defaultVisibleCols)
            {
                if (_dgvSDS.Columns.Contains(defCol) && _dgvSDS.Columns[defCol].Visible)
                {
                    _dgvSDS.Columns[defCol].DisplayIndex = currentIndex++;
                }
            }
            
            // 2. 排序其他由使用者手動勾選擴充顯示的欄位
            foreach (DataGridViewColumn col in _dgvSDS.Columns)
            {
                if (col.Visible && !_defaultVisibleCols.Contains(col.Name) && !col.Name.Equals("Id", StringComparison.OrdinalIgnoreCase))
                {
                    col.DisplayIndex = currentIndex++;
                }
            }

            if (isFirstTime) SaveVisibilitySettings();
        }

        private void OpenColumnSettings()
        {
            try
            {
                List<string> allCols = new List<string>();
                if (_dgvSDS.Columns.Count > 0) {
                    foreach(DataGridViewColumn c in _dgvSDS.Columns) {
                        allCols.Add(c.Name);
                    }
                }

                if (allCols.Count == 0)
                {
                    MessageBox.Show("目前找不到資料表，請先匯入資料。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (Form f = new Form())
                {
                    f.Text = "⚙️ 看板顯示欄位設定";
                    f.Size = new Size(380, 550);
                    f.StartPosition = FormStartPosition.CenterParent;
                    f.FormBorderStyle = FormBorderStyle.FixedDialog;
                    f.MaximizeBox = false; 
                    f.MinimizeBox = false;
                    f.BackColor = Color.White;

                    Label lbl = new Label { 
                        Text = "請勾選要在看板與報表中顯示的項目：", 
                        Dock = DockStyle.Top, 
                        Height = 50, 
                        Padding = new Padding(10), 
                        Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                        ForeColor = Color.SteelBlue
                    };
                    
                    CheckedListBox clb = new CheckedListBox { 
                        Dock = DockStyle.Fill, 
                        CheckOnClick = true, 
                        Font = new Font("Microsoft JhengHei UI", 11F),
                        BorderStyle = BorderStyle.None,
                        BackColor = Color.FromArgb(250, 250, 250)
                    };
                    
                    foreach (var colName in allCols)
                    {
                        if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase)) continue;
                        
                        // 取得目前的勾選狀態
                        bool isChecked = _columnVisibility.ContainsKey(colName) ? _columnVisibility[colName] : _defaultVisibleCols.Contains(colName);
                        clb.Items.Add(colName, isChecked);
                    }

                    Button btnSave = new Button { 
                        Text = "💾 儲存設定並套用", 
                        Dock = DockStyle.Bottom, 
                        Height = 55, 
                        BackColor = Color.ForestGreen, 
                        ForeColor = Color.White, 
                        Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                        Cursor = Cursors.Hand,
                        FlatStyle = FlatStyle.Flat
                    };
                    
                    btnSave.Click += (s, e) => {
                        _columnVisibility.Clear();
                        for (int i = 0; i < clb.Items.Count; i++)
                        {
                            _columnVisibility[clb.Items[i].ToString()] = clb.GetItemChecked(i);
                        }
                        SaveVisibilitySettings();
                        ApplyVisibility();
                        f.DialogResult = DialogResult.OK;
                    };

                    f.Controls.Add(clb);
                    f.Controls.Add(lbl);
                    f.Controls.Add(btnSave);
                    f.ShowDialog();
                }
            }
            catch (Exception ex) { MessageBox.Show("開啟設定視窗失敗：" + ex.Message); }
        }

        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            if (File.Exists(VisibilityFile))
            {
                try
                {
                    string[] lines = File.ReadAllLines(VisibilityFile, Encoding.UTF8);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 2)
                        {
                            _columnVisibility[parts[0]] = (parts[1] == "1");
                        }
                    }
                }
                catch { }
            }
        }

        private void SaveVisibilitySettings()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (var kvp in _columnVisibility)
                {
                    sb.AppendLine($"{kvp.Key}|{(kvp.Value ? "1" : "0")}");
                }
                File.WriteAllText(VisibilityFile, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        // =========================================================================
        // 導出 A4 危害性化學品清單 PDF 功能 (直式，每筆一頁)
        // =========================================================================
        private void ExportToHazardousListPdfDirectly()
        {
            if (_dgvSDS.DataSource == null || _dgvSDS.Rows.Count == 0)
            {
                MessageBox.Show("目前沒有數據可供導出。"); return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "危害性化學品清單_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                    
                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = false; 
                    pd.DefaultPageSettings.Margins = new Margins(50, 50, 60, 60);
                    
                    int currentRow = 0;

                    pd.PrintPage += (s, e) => {
                        Graphics g = e.Graphics;
                        float x = e.MarginBounds.Left;
                        float y = e.MarginBounds.Top;
                        float w = e.MarginBounds.Width;

                        Font fTitle = new Font("DFKai-SB", 22F, FontStyle.Bold); 
                        Font fBody = new Font("DFKai-SB", 14F);
                        Font fSmall = new Font("DFKai-SB", 12F);

                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        string separator = "※※※※※※※※※※※※※※※※※※※※※※※※※";

                        DataGridViewRow row = _dgvSDS.Rows[currentRow];
                        
                        string GetVal(string colName) {
                            if (_dgvSDS.Columns.Contains(colName) && row.Cells[colName].Value != null)
                                return row.Cells[colName].Value.ToString();
                            return "";
                        }

                        g.DrawString("危害性化學品清單", fTitle, Brushes.Black, new RectangleF(x, y, w, 40), sfCenter);
                        y += 60;

                        g.DrawString(separator, fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString($"化學品名稱：{GetVal("化學物質名稱")}", fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString($"其他名稱：{GetVal("其它化學物質名稱")}", fBody, Brushes.Black, x, y); y += 30;
                        
                        g.DrawString($"安全資料表索引碼：{GetVal("廠內編號")}", fBody, Brushes.Black, x, y); y += 30;
                        
                        g.DrawString(separator, fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString($"製造者、輸入者或供應者：{GetVal("供應商")}", fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString($"供應商地址：{GetVal("供應商地址")}", fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString($"供應商電話：{GetVal("供應商電話")}", fBody, Brushes.Black, x, y); y += 30;

                        g.DrawString(separator, fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString("使用資料", fBody, Brushes.Black, x, y); y += 30;

                        float col1X = x;
                        float col2X = x + 160;
                        float col3X = x + 340;
                        float col4X = x + 520;

                        g.DrawString("地  點", fBody, Brushes.Black, col1X, y);
                        g.DrawString("平均數量", fBody, Brushes.Black, col2X, y);
                        g.DrawString("最大數量", fBody, Brushes.Black, col3X, y);
                        g.DrawString("使用者", fBody, Brushes.Black, col4X, y);
                        y += 35;

                        g.DrawString(GetVal("使用地點"), fBody, Brushes.Black, col1X, y);
                        g.DrawString(GetVal("使用平均量"), fBody, Brushes.Black, col2X, y);
                        g.DrawString(GetVal("使用最大量"), fBody, Brushes.Black, col3X, y);
                        g.DrawString(GetVal("使用單位"), fBody, Brushes.Black, col4X, y);
                        y += 35;

                        for (int i = 0; i < 2; i++) {
                            g.DrawLine(Pens.Black, col1X, y + 20, col1X + 100, y + 20);
                            g.DrawLine(Pens.Black, col2X, y + 20, col2X + 100, y + 20);
                            g.DrawLine(Pens.Black, col3X, y + 20, col3X + 100, y + 20);
                            g.DrawLine(Pens.Black, col4X, y + 20, col4X + 100, y + 20);
                            y += 35;
                        }

                        g.DrawString(separator, fBody, Brushes.Black, x, y); y += 30;
                        g.DrawString("貯存資料", fBody, Brushes.Black, x, y); y += 30;

                        g.DrawString("地  點", fBody, Brushes.Black, col1X, y);
                        g.DrawString("平均數量", fBody, Brushes.Black, col2X, y);
                        g.DrawString("最大數量", fBody, Brushes.Black, col3X, y);
                        y += 35;

                        g.DrawString(GetVal("貯存地點"), fBody, Brushes.Black, col1X, y);
                        g.DrawString(GetVal("平均貯存量"), fBody, Brushes.Black, col2X, y);
                        g.DrawString(GetVal("最大貯存量"), fBody, Brushes.Black, col3X, y);
                        y += 35;

                        for (int i = 0; i < 2; i++) {
                            g.DrawLine(Pens.Black, col1X, y + 20, col1X + 100, y + 20);
                            g.DrawLine(Pens.Black, col2X, y + 20, col2X + 100, y + 20);
                            g.DrawLine(Pens.Black, col3X, y + 20, col3X + 100, y + 20);
                            y += 35;
                        }

                        g.DrawString(separator, fBody, Brushes.Black, x, y); y += 40;
                        string dateStr = $"製單日期：{DateTime.Now.Year} 年 {DateTime.Now.Month:D2} 月 {DateTime.Now.Day:D2} 日";
                        g.DrawString(dateStr, fBody, Brushes.Black, x, y);

                        g.DrawString("8-ES-B09-01", fSmall, Brushes.Black, x, e.MarginBounds.Bottom - 20);

                        currentRow++;
                        if (currentRow < _dgvSDS.Rows.Count) {
                            e.HasMorePages = true;
                        } else {
                            e.HasMorePages = false;
                            currentRow = 0; 
                        }
                    };

                    try {
                        pd.Print();
                        MessageBox.Show("危害性化學品清單 PDF 匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void ExportToPdf()
        {
            if (_dgvSDS.DataSource == null || _dgvSDS.Rows.Count == 0)
            {
                MessageBox.Show("目前沒有數據可供導出。");
                return;
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; 
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int rowIndex = 0;

            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;
                float pageWidth = e.MarginBounds.Width;

                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fSubTitle = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold); 
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

                StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                StringFormat sfBody = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                string mainTitle = "台灣玻璃工業股份有限公司 - 彰濱廠";
                string subTitle = "化學品清單一覽表";
                string pageInfo = $"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   頁碼：{rowIndex / 20 + 1}";

                g.DrawString(mainTitle, fTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 40), sfCenter); y += 40;
                g.DrawString(subTitle, fSubTitle, Brushes.Black, new RectangleF(x, y, pageWidth, 35), sfCenter); y += 35;
                g.DrawString(pageInfo, fBody, Brushes.Gray, new RectangleF(x, y, pageWidth, 30), sfCenter); y += 30;

                var visCols = _dgvSDS.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();
                if (visCols.Count == 0) return;

                float totalGridWidth = visCols.Sum(c => c.Width);
                float scale = e.MarginBounds.Width / totalGridWidth;
                if (scale > 1.2f) scale = 1.2f; 

                float currX = x;
                float rowH = 32;
                foreach (var col in visCols)
                {
                    RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                    g.FillRectangle(Brushes.DimGray, rect);
                    g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                    g.DrawString(col.HeaderText, fHead, Brushes.White, rect, sfCenter);
                    currX += col.Width * scale;
                }
                y += rowH;

                while (rowIndex < _dgvSDS.Rows.Count)
                {
                    currX = x;
                    foreach (var col in visCols)
                    {
                        RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                        g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                        string val = _dgvSDS[col.Index, rowIndex].Value?.ToString() ?? "";
                        g.DrawString(val, fBody, Brushes.Black, rect, sfBody);
                        currX += col.Width * scale;
                    }
                    y += rowH;
                    rowIndex++;

                    if (y + rowH > e.MarginBounds.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }
                }
                
                e.HasMorePages = false;
                rowIndex = 0; 
            };

            PrintPreviewDialog ppd = new PrintPreviewDialog { 
                Document = pd, 
                Width = 1024, 
                Height = 768, 
                WindowState = FormWindowState.Maximized,
                UseAntiAlias = true
            };
            ppd.ShowDialog();
        }
    }
}
