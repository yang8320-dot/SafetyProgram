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
        private DataGridView _dgvSDS;
        private const string DbName = "Chemical";
        private const string TableName = "SDS_Inventory";
        
        // 設定檔路徑
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChemSDS_Visibility.txt");
        
        // 記憶欄位顯示狀態
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        public Control GetView()
        {
            LoadVisibilitySettings(); // 載入記憶的設定

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 第一行：功能按鈕區
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            Button btnPdf = new Button { Text = "📤 導出清冊 PDF", Size = new Size(180, 40), BackColor = Color.DarkCyan, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.Click += (s, e) => ExportToPdf();

            // 🟢 新增：設定顯示欄位按鈕
            Button btnSettings = new Button { Text = "⚙️ 設定顯示欄位", Size = new Size(180, 40), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSettings.Click += (s, e) => OpenColumnSettings();

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(btnSettings);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // 第二大框
            GroupBox boxMain = new GroupBox { Text = "📋 化學品管理綜合看板", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Padding = new Padding(15) };
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F)); 
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // 小框 1：標題區
            Panel pnlTitle = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 245, 250), Margin = new Padding(0,0,0,15) };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Solid);
            
            Label lblCompany = new Label { Text = "台灣玻璃工業股份有限公司 - 彰濱廠", Dock = DockStyle.Top, Height = 45, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.MidnightBlue };
            Label lblSubTitle = new Label { Text = "化學品清單一覽表 (SDS 定期追蹤管理)", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.FromArgb(60, 60, 60) };
            pnlTitle.Controls.Add(lblSubTitle);
            pnlTitle.Controls.Add(lblCompany);

            // 小框 2：SDS 數據表格
            GroupBox boxGrid = new GroupBox { Text = "SDS 安全資料表庫存清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
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
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            _dgvSDS.DataSource = dt;
            ApplyVisibility(); // 套用顯示/隱藏邏輯
        }

        private void ApplyVisibility()
        {
            foreach (DataGridViewColumn col in _dgvSDS.Columns)
            {
                if (col.Name == "Id") { col.Visible = false; continue; }
                
                // 根據記憶字典決定顯示與否，若字典沒資料則預設顯示 (除了附件)
                if (_columnVisibility.ContainsKey(col.Name))
                {
                    col.Visible = _columnVisibility[col.Name];
                }
                else
                {
                    col.Visible = (col.Name != "附件檔案");
                }
            }
        }

        private void OpenColumnSettings()
        {
            // 取得目前資料庫所有欄位
            List<string> allCols = DataManager.GetColumnNames(DbName, TableName);

            using (Form f = new Form())
            {
                f.Text = "設定顯示欄位";
                f.Size = new Size(350, 500);
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false; f.MinimizeBox = false;

                Label lbl = new Label { Text = "請勾選欲顯示在看板上的欄位：", Dock = DockStyle.Top, Height = 40, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                
                foreach (var colName in allCols)
                {
                    if (colName == "Id") continue;
                    bool isChecked = true; // 預設勾選
                    if (_columnVisibility.ContainsKey(colName)) isChecked = _columnVisibility[colName];
                    else if (colName == "附件檔案") isChecked = false; // 附件預設不顯在看板

                    clb.Items.Add(colName, isChecked);
                }

                Button btnSave = new Button { Text = "💾 儲存並套用", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnSave.Click += (s, e) => {
                    _columnVisibility.Clear();
                    for (int i = 0; i < clb.Items.Count; i++)
                    {
                        _columnVisibility[clb.Items[i].ToString()] = clb.GetItemChecked(i);
                    }
                    SaveVisibilitySettings();
                    ApplyVisibility();
                    f.Close();
                };

                f.Controls.Add(clb);
                f.Controls.Add(lbl);
                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
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
                            _columnVisibility[parts[0]] = parts[1] == "1";
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
                List<string> lines = new List<string>();
                foreach (var kvp in _columnVisibility)
                {
                    lines.Add($"{kvp.Key}|{(kvp.Value ? "1" : "0")}");
                }
                File.WriteAllLines(VisibilityFile, lines, Encoding.UTF8);
            }
            catch { }
        }

        private void ExportToPdf()
        {
            if (_dgvSDS.Rows.Count == 0) return;

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true;
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            int rowIndex = 0;

            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), Brushes.MidnightBlue, x, y);
                y += 35;
                g.DrawString("化學品清單一覽表 (SDS)", new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Brushes.Black, x, y);
                y += 30;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}", new Font("Microsoft JhengHei UI", 10F), Brushes.Gray, x, y);
                y += 25;

                // 🟢 只抓取目前「可見」的欄位進行 PDF 列印
                var visCols = _dgvSDS.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                float totalW = visCols.Sum(c => c.Width);
                float scale = e.MarginBounds.Width / totalW;
                if (scale > 1) scale = 1;

                float currX = x;
                foreach (var col in visCols) {
                    RectangleF rect = new RectangleF(currX, y, col.Width * scale, 30);
                    g.FillRectangle(Brushes.DimGray, rect);
                    g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                    g.DrawString(col.HeaderText, new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold), Brushes.White, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    currX += col.Width * scale;
                }
                y += 30;

                while (rowIndex < _dgvSDS.Rows.Count) {
                    currX = x;
                    foreach (var col in visCols) {
                        RectangleF rect = new RectangleF(currX, y, col.Width * scale, 30);
                        g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                        string val = _dgvSDS[col.Index, rowIndex].Value?.ToString() ?? "";
                        g.DrawString(val, new Font("Microsoft JhengHei UI", 9F), Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
                        currX += col.Width * scale;
                    }
                    y += 30;
                    rowIndex++;
                    if (y + 30 > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                }
                e.HasMorePages = false;
                rowIndex = 0;
            };

            PrintPreviewDialog ppd = new PrintPreviewDialog { Document = pd, Width = 1024, Height = 768, WindowState = FormWindowState.Maximized };
            ppd.ShowDialog();
        }
    }
}
