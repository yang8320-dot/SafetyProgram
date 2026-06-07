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
        
        // 確保新的 10 個預設欄位直接生效
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChemSDS_Columns_v4.txt");
        
        // 用於儲存欄位順序與可見性的結構
        private class ColConfig {
            public string Name { get; set; }
            public bool IsVisible { get; set; }
        }
        private List<ColConfig> _columnSettings = new List<ColConfig>();

        // 您指定的 10 個預設顯示欄位與排序 (對應真實欄位名稱)
        private readonly string[] _defaultVisibleCols = { "項次", "化學物質名稱", "危害標示", "供應商", "供應商電話", "使用單位", "貯存地點", "使用最大量", "SDS版本日期", "備註" };

        public Control GetView()
        {
            // 防呆：確保資料表已建立，避免因為空表導致讀取崩潰畫面擠壓
            string schema = TableSchemaManager.SchemaMap.ContainsKey(TableName) ? TableSchemaManager.SchemaMap[TableName] : "[日期] TEXT, [備註] TEXT";
            DataManager.InitTable(DbName, TableName, $"CREATE TABLE IF NOT EXISTS [{TableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});");

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
                Dock = DockStyle.Top, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 15),
                WrapContents = false
            };
            
            Button btnPdf = new Button { 
                Text = "📤 導出 SDS 清冊", 
                Size = new Size(230, 45), 
                BackColor = Color.DarkCyan, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 15, 0)
            };
            btnPdf.FlatAppearance.BorderSize = 0;
            btnPdf.Click += (s, e) => ExportToPdf();

            Button btnHazardousPdf = new Button { 
                Text = "📄 導出 危害性化學品清單", 
                Size = new Size(280, 45), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 15, 0)
            };
            btnHazardousPdf.FlatAppearance.BorderSize = 0;
            btnHazardousPdf.Click += (s, e) => ExportToHazardousListPdfDirectly();

            Button btnSettings = new Button { 
                Text = "⚙️ 設定顯示欄位與排序", 
                Size = new Size(240, 45), 
                BackColor = Color.LightSlateGray, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand, 
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 0, 0) 
            };
            btnSettings.FlatAppearance.BorderSize = 0;
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
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F)); 
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  

            Panel pnlTitle = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.FromArgb(240, 245, 250), 
                Margin = new Padding(0, 0, 0, 15) 
            };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.LightSkyBlue, ButtonBorderStyle.Solid);
            
            Label lblSubTitle = new Label { 
                Text = "化學品清單一覽表", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                ForeColor = Color.SteelBlue 
            };
            pnlTitle.Controls.Add(lblSubTitle);

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
                AutoGenerateColumns = false, 
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
            
            // 🟢 核心攔截：儲存格自訂繪製 (純圖示處理)
            _dgvSDS.CellPainting += DgvSDS_CellPainting;

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
                    if (!dt.Columns.Contains("項次"))
                    {
                        DataColumn seqCol = new DataColumn("項次", typeof(int));
                        dt.Columns.Add(seqCol);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dt.Rows[i]["項次"] = i + 1;
                        }
                    }

                    ApplyVisibilityAndGenerateColumns(dt); 
                    
                    _dgvSDS.DataSource = dt;
                    
                    if (_dgvSDS.Columns.Contains("項次")) {
                        _dgvSDS.Columns["項次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        _dgvSDS.Columns["項次"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        _dgvSDS.Columns["項次"].Width = 60;
                    }
                }
                else
                {
                    _dgvSDS.DataSource = null;
                }
            }
            catch { _dgvSDS.DataSource = null; }
        }

        private void ApplyVisibilityAndGenerateColumns(DataTable dt)
        {
            if (dt == null) return;

            List<string> actualCols = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            if (_columnSettings.Count == 0)
            {
                foreach (string defCol in _defaultVisibleCols)
                {
                    if (actualCols.Contains(defCol))
                    {
                        _columnSettings.Add(new ColConfig { Name = defCol, IsVisible = true });
                    }
                }

                foreach (string colName in actualCols)
                {
                    if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!_columnSettings.Any(c => c.Name == colName))
                    {
                        _columnSettings.Add(new ColConfig { Name = colName, IsVisible = false });
                    }
                }
                SaveVisibilitySettings();
            }
            else
            {
                foreach (string colName in actualCols)
                {
                    if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!_columnSettings.Any(c => c.Name == colName))
                    {
                        _columnSettings.Add(new ColConfig { Name = colName, IsVisible = false });
                    }
                }
                _columnSettings.RemoveAll(c => !actualCols.Contains(c.Name));
            }

            _dgvSDS.Columns.Clear();
            int currentDisplayIndex = 0;

            foreach (var cfg in _columnSettings)
            {
                if (actualCols.Contains(cfg.Name))
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn
                    {
                        Name = cfg.Name,
                        HeaderText = cfg.Name,
                        DataPropertyName = cfg.Name, 
                        Visible = cfg.IsVisible
                    };
                    
                    _dgvSDS.Columns.Add(col);

                    if (cfg.IsVisible)
                    {
                        col.DisplayIndex = currentDisplayIndex++;
                    }
                }
            }
        }

        private void OpenColumnSettings()
        {
            try
            {
                if (_columnSettings.Count == 0)
                {
                    MessageBox.Show("目前找不到資料庫資料，請先匯入資料以建立結構。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (Form f = new Form())
                {
                    f.Text = "⚙️ 顯示欄位與排序設定";
                    f.Size = new Size(500, 600);
                    f.StartPosition = FormStartPosition.CenterParent;
                    f.FormBorderStyle = FormBorderStyle.FixedDialog;
                    f.MaximizeBox = false; 
                    f.MinimizeBox = false;
                    f.BackColor = Color.White;

                    Label lbl = new Label { 
                        Text = "勾選顯示項目，並可透過右側按鈕自訂欄位左右排列順序：", 
                        Dock = DockStyle.Top, 
                        Height = 50, 
                        Padding = new Padding(10), 
                        Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                        ForeColor = Color.SteelBlue
                    };
                    
                    Panel pnlCenter = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                    
                    CheckedListBox clb = new CheckedListBox { 
                        Dock = DockStyle.Fill, 
                        CheckOnClick = true, 
                        Font = new Font("Microsoft JhengHei UI", 12F),
                        BorderStyle = BorderStyle.FixedSingle,
                        BackColor = Color.FromArgb(250, 250, 250)
                    };

                    foreach (var cfg in _columnSettings)
                    {
                        clb.Items.Add(cfg.Name, cfg.IsVisible);
                    }

                    Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 120, Padding = new Padding(10, 0, 0, 0) };
                    
                    Button btnUp = new Button { Text = "↑ 上移", Width = 100, Height = 40, Margin = new Padding(0, 0, 0, 10), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.WhiteSmoke };
                    btnUp.Click += (s, e) => {
                        if (clb.SelectedIndex > 0) {
                            int idx = clb.SelectedIndex;
                            object item = clb.SelectedItem;
                            bool isChecked = clb.GetItemChecked(idx);
                            
                            clb.Items.RemoveAt(idx);
                            clb.Items.Insert(idx - 1, item);
                            clb.SetItemChecked(idx - 1, isChecked);
                            clb.SelectedIndex = idx - 1;
                        }
                    };

                    Button btnDown = new Button { Text = "↓ 下移", Width = 100, Height = 40, Top = 50, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.WhiteSmoke };
                    btnDown.Click += (s, e) => {
                        if (clb.SelectedIndex < clb.Items.Count - 1 && clb.SelectedIndex != -1) {
                            int idx = clb.SelectedIndex;
                            object item = clb.SelectedItem;
                            bool isChecked = clb.GetItemChecked(idx);
                            
                            clb.Items.RemoveAt(idx);
                            clb.Items.Insert(idx + 1, item);
                            clb.SetItemChecked(idx + 1, isChecked);
                            clb.SelectedIndex = idx + 1;
                        }
                    };

                    pnlRight.Controls.Add(btnUp);
                    pnlRight.Controls.Add(btnDown);

                    pnlCenter.Controls.Add(clb);
                    pnlCenter.Controls.Add(pnlRight);

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
                        _columnSettings.Clear();
                        for (int i = 0; i < clb.Items.Count; i++)
                        {
                            _columnSettings.Add(new ColConfig {
                                Name = clb.Items[i].ToString(),
                                IsVisible = clb.GetItemChecked(i)
                            });
                        }
                        SaveVisibilitySettings();
                        
                        LoadData();

                        f.DialogResult = DialogResult.OK;
                    };

                    f.Controls.Add(pnlCenter);
                    f.Controls.Add(lbl);
                    f.Controls.Add(btnSave);
                    f.ShowDialog();
                }
            }
            catch (Exception ex) { MessageBox.Show("開啟設定視窗失敗：" + ex.Message); }
        }

        private void LoadVisibilitySettings()
        {
            _columnSettings.Clear();
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
                            _columnSettings.Add(new ColConfig {
                                Name = parts[0],
                                IsVisible = (parts[1] == "1")
                            });
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
                foreach (var cfg in _columnSettings)
                {
                    sb.AppendLine($"{cfg.Name}|{(cfg.IsVisible ? "1" : "0")}");
                }
                File.WriteAllText(VisibilityFile, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        // =========================================================================
        // 🟢 導出 A4 危害性化學品清單 PDF 功能 (套用圖示顯示)
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
                    Form activeForm = Form.ActiveForm;
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;
                    
                    try
                    {
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

                            Font fTitle = new Font("DFKai-SB", 26F, FontStyle.Bold); 
                            Font fBody = new Font("DFKai-SB", 14F);
                            Font fSmall = new Font("DFKai-SB", 12F);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                            string separator = "※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※";

                            DataGridViewRow row = _dgvSDS.Rows[currentRow];
                            
                            string GetVal(string colName) {
                                if (_dgvSDS.Columns.Contains(colName) && row.Cells[colName].Value != null)
                                    return row.Cells[colName].Value.ToString();
                                return "";
                            }

                            // 1. 唯一標題：危害性化學品清單
                            g.DrawString("危害性化學品清單", fTitle, Brushes.Black, new RectangleF(x, y, w, 50), sfCenter);
                            y += 80;

                            // 以下為實體內容
                            g.DrawString(separator, fBody, Brushes.Black, x, y); y += 30;
                            g.DrawString($"化學品名稱：{GetVal("化學物質名稱")}", fBody, Brushes.Black, x, y); y += 30;
                            g.DrawString($"其他名稱：{GetVal("其它化學物質名稱")}", fBody, Brushes.Black, x, y); y += 30;
                            g.DrawString($"安全資料表索引碼：{GetVal("廠內編號")}", fBody, Brushes.Black, x, y); y += 30;
                            
                            // 🟢 危害標示 (處理圖示)
                            g.DrawString($"危害標示：", fBody, Brushes.Black, x, y);
                            string hazardVal = GetVal("危害標示");
                            var hazardIcons = PdfHelper.GetIconsFromCache(TableName, "危害標示", hazardVal);
                            
                            if (hazardIcons.Count > 0) {
                                float imgStartX = x + 100;
                                foreach (var img in hazardIcons) {
                                    g.DrawImage(img, imgStartX, y - 4, 24, 24);
                                    imgStartX += 28;
                                }
                            } else {
                                g.DrawString(hazardVal, fBody, Brushes.Black, x + 100, y);
                            }
                            y += 30;

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

                            // 底部表單號 (無頁碼)
                            g.DrawString("8-ES-B09-01 危害性化學品清單", fSmall, Brushes.Black, x, e.MarginBounds.Bottom - 20);

                            currentRow++;
                            if (currentRow < _dgvSDS.Rows.Count) {
                                e.HasMorePages = true;
                            } else {
                                e.HasMorePages = false;
                                currentRow = 0; 
                            }
                        };

                        pd.Print();
                        MessageBox.Show("危害性化學品清單 PDF 匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        // =========================================================================
        // 導出 SDS 清冊 PDF 功能 (套用圖示顯示)
        // =========================================================================
        private void ExportToPdf()
        {
            if (_dgvSDS.DataSource == null || _dgvSDS.Rows.Count == 0)
            {
                MessageBox.Show("目前沒有數據可供導出。");
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "SDS清冊_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    Form activeForm = Form.ActiveForm;
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;

                    try
                    {
                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = true; 
                        pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
                        
                        int rowIndex = 0;
                        int pageNumber = 1;

                        var visCols = _dgvSDS.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();
                        if (visCols.Count == 0) return;

                        // ====== 先計算總頁數 ======
                        int totalPages = 1;
                        float simH = 827f - 60f;  
                        
                        float simStartY = 30f + 40f + 40f + 35f + 30f + 32f; 
                        float simCurrentY = simStartY;
                        float simBottomLimit = simH - 30f; 

                        for (int i = 0; i < _dgvSDS.Rows.Count; i++) {
                            float rowH = 32f; 
                            if (simCurrentY + rowH > simBottomLimit) {
                                totalPages++;
                                simCurrentY = simStartY + rowH;
                            } else {
                                simCurrentY += rowH;
                            }
                        }

                        // ====== 正式繪製 ======
                        pd.PrintPage += (s, e) => {
                            Graphics g = e.Graphics;
                            float x = e.MarginBounds.Left;
                            float y = e.MarginBounds.Top;
                            float w = e.MarginBounds.Width;

                            Font fTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                            Font fSubTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold); 
                            Font fSign = new Font("Microsoft JhengHei UI", 12F);
                            Font fDate = new Font("Microsoft JhengHei UI", 11F);
                            Font fBody = new Font("Microsoft JhengHei UI", 9F);
                            Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                            StringFormat sfBody = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                            // 1. 第一行：大標題置中
                            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                            y += 40;

                            // 2. 第二行：副標題置中
                            g.DrawString("化學品清單一覽表", fSubTitle, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                            y += 40;

                            // 3. 第三行：簽核置中
                            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
                            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                            y += 35;

                            // 4. 第四行：查詢條件留空高度
                            g.DrawString("", fDate, Brushes.DimGray, new RectangleF(x, y, w, 20), new StringFormat { Alignment = StringAlignment.Near }); 
                            y += 30;

                            float totalGridWidth = visCols.Sum(c => c.Width);
                            float scale = w / totalGridWidth;
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
                                    
                                    // 🟢 檢查是否有圖示
                                    var icons = PdfHelper.GetIconsFromCache(TableName, col.Name, val);
                                    if (icons.Count > 0)
                                    {
                                        float imgSize = 18f; // PDF上的圖示大小
                                        float imgY = y + (rowH - imgSize) / 2;
                                        float startX = currX + 4;
                                        foreach (var img in icons)
                                        {
                                            g.DrawImage(img, startX, imgY, imgSize, imgSize);
                                            startX += imgSize + 4;
                                        }
                                    }
                                    else
                                    {
                                        // 沒圖示就畫純文字
                                        RectangleF textRect = new RectangleF(rect.X + 2, rect.Y, rect.Width - 4, rect.Height);
                                        g.DrawString(val, fBody, Brushes.Black, textRect, sfBody);
                                    }
                                    
                                    currX += col.Width * scale;
                                }
                                y += rowH;
                                rowIndex++;

                                if (y + rowH > e.MarginBounds.Bottom - 30)
                                {
                                    break;
                                }
                            }
                            
                            // 底部標註與頁碼
                            g.DrawString("8-ES-B09-01 化學品清冊一覽表", fDate, Brushes.Black, x, e.MarginBounds.Bottom - 15);
                            g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fDate, Brushes.Black, new RectangleF(x, e.MarginBounds.Bottom - 15, w, 20), sfCenter);

                            if (rowIndex < _dgvSDS.Rows.Count) {
                                pageNumber++;
                                e.HasMorePages = true;
                            } else {
                                e.HasMorePages = false;
                                rowIndex = 0; 
                                pageNumber = 1;
                            }
                        };

                        pd.Print();
                        MessageBox.Show("SDS 清冊 PDF 匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        // =========================================================================
        // 🟢 讓化學品看板的 Grid 也支援純圖示/多圖示並排顯示
        // =========================================================================
        private void DgvSDS_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string colName = _dgvSDS.Columns[e.ColumnIndex].Name;

            if (e.Value != null)
            {
                var icons = PdfHelper.GetIconsFromCache(TableName, colName, e.Value.ToString());
                if (icons.Count > 0)
                {
                    // 🟢 核心修正：只畫背景與邊框，刻意排除文字
                    e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border | DataGridViewPaintParts.Focus | DataGridViewPaintParts.SelectionBackground);
                    
                    int imgSize = 24; 
                    int startX = e.CellBounds.X + 6;
                    int imgY = e.CellBounds.Y + (e.CellBounds.Height - imgSize) / 2;

                    foreach (var img in icons)
                    {
                        e.Graphics.DrawImage(img, startX, imgY, imgSize, imgSize);
                        startX += imgSize + 4;
                    }
                    e.Handled = true; // 告知系統已處理完成
                }
            }
        }
    }
}
