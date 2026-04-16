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
        
        // 記憶設定檔路徑
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChemSDS_Visibility.txt");
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        public Control GetView()
        {
            LoadVisibilitySettings(); // 啟動時先載入偏好設定

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 第一行：功能按鈕區 (導出與設定)
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            Button btnPdf = new Button { Text = "📤 導出清冊 PDF", Size = new Size(200, 40), BackColor = Color.DarkCyan, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.Click += (s, e) => ExportToPdf();

            Button btnSettings = new Button { Text = "⚙️ 設定顯示欄位", Size = new Size(200, 40), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSettings.Click += (s, e) => OpenColumnSettings();

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(btnSettings);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // 第二大框：包含標題區與數據區
            GroupBox boxMain = new GroupBox { Text = "📋 化學品管理綜合看板", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Padding = new Padding(15) };
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F)); // 小框1: 標題
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 小框2: 數據

            // 小框 1：企業標題與報表名稱
            Panel pnlTitle = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 245, 250), Margin = new Padding(0,0,0,15) };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Solid);
            
            Label lblCompany = new Label { Text = "台灣玻璃工業股份有限公司 - 彰濱廠", Dock = DockStyle.Top, Height = 45, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.MidnightBlue };
            Label lblSubTitle = new Label { Text = "化學品清單一覽表 (SDS 定期追蹤管理)", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.FromArgb(60, 60, 60) };
            pnlTitle.Controls.Add(lblSubTitle);
            pnlTitle.Controls.Add(lblCompany);

            // 小框 2：數據顯示區 (SDS 資料表)
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
            try
            {
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                if (dt != null)
                {
                    _dgvSDS.DataSource = dt;
                    ApplyVisibility();
                }
            }
            catch { _dgvSDS.DataSource = null; } // 找不到資料表時防空處理
        }

        private void ApplyVisibility()
        {
            if (_dgvSDS.Columns.Count == 0) return;
            foreach (DataGridViewColumn col in _dgvSDS.Columns)
            {
                if (col.Name == "Id") { col.Visible = false; continue; }
                if (_columnVisibility.ContainsKey(col.Name))
                {
                    col.Visible = _columnVisibility[col.Name];
                }
                else
                {
                    // 預設排除「附件檔案」路徑字串顯示
                    col.Visible = (col.Name != "附件檔案");
                }
            }
        }

        private void OpenColumnSettings()
        {
            List<string> allCols = DataManager.GetColumnNames(DbName, TableName);
            if (allCols == null || allCols.Count == 0)
            {
                MessageBox.Show("目前找不到資料表，請確認資料庫是否已初始化。", "提示");
                return;
            }

            using (Form f = new Form { Text = "⚙️ 設定看板顯示欄位", Size = new Size(350, 520), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲在看板顯示的數據項目：", Dock = DockStyle.Top, Height = 40, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 11F), BorderStyle = BorderStyle.None, BackColor = Color.WhiteSmoke };
                
                foreach (var colName in allCols)
                {
                    if (colName == "Id") continue;
                    bool isChecked = _columnVisibility.ContainsKey(colName) ? _columnVisibility[colName] : (colName != "附件檔案");
                    clb.Items.Add(colName, isChecked);
                }

                Button btnOk = new Button { Text = "💾 儲存設定並套用", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnOk.Click += (s, e) => {
                    _columnVisibility.Clear();
                    for (int i = 0; i < clb.Items.Count; i++) _columnVisibility[clb.Items[i].ToString()] = clb.GetItemChecked(i);
                    SaveVisibilitySettings();
                    ApplyVisibility();
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(clb);
                f.Controls.Add(lbl);
                f.Controls.Add(btnOk);
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
                        if (parts.Length == 2) _columnVisibility[parts[0]] = (parts[1] == "1");
                    }
                }
                catch { }
            }
        }

        private void SaveVisibilitySettings()
        {
            try
            {
                List<string> lines = _columnVisibility.Select(kvp => $"{kvp.Key}|{(kvp.Value ? "1" : "0")}").ToList();
                File.WriteAllLines(VisibilityFile, lines, Encoding.UTF8);
            }
            catch { }
        }

        private void ExportToPdf()
        {
            if (_dgvSDS.Rows.Count == 0) { MessageBox.Show("無資料可供導出。"); return; }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; // A4 橫向
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int rowIndex = 0;
            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fSub = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                // 報表標題區
                g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.MidnightBlue, x, y);
                y += 40;
                g.DrawString("化學品清單一覽表 (SDS)", fSub, Brushes.Black, x, y);
                y += 35;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   頁碼：{rowIndex / 20 + 1}", fBody, Brushes.Gray, x, y);
                y += 30;

                // 篩選出目前在看板上「被勾選為顯示」的欄位
                var visCols = _dgvSDS.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                if (visCols.Count == 0) return;

                float totalW = visCols.Sum(c => c.Width);
                float scale = e.MarginBounds.Width / totalW;
                if (scale > 1.2f) scale = 1.2f; // 防止過度拉大

                // 畫表格標頭
                float currX = x;
                float rowH = 30;
                foreach (var col in visCols) {
                    RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                    g.FillRectangle(Brushes.DimGray, rect);
                    g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                    g.DrawString(col.HeaderText, fHead, Brushes.White, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    currX += col.Width * scale;
                }
                y += rowH;

                // 畫表格內容 (含分頁邏輯)
                while (rowIndex < _dgvSDS.Rows.Count) {
                    currX = x;
                    foreach (var col in visCols) {
                        RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                        g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                        string val = _dgvSDS[col.Index, rowIndex].Value?.ToString() ?? "";
                        g.DrawString(val, fBody, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
                        currX += col.Width * scale;
                    }
                    y += rowH;
                    rowIndex++;

                    // A4 分頁高度檢查
                    if (y + rowH > e.MarginBounds.Bottom) {
                        e.HasMorePages = true;
                        return;
                    }
                }
                e.HasMorePages = false;
                rowIndex = 0; 
            };

            PrintPreviewDialog ppd = new PrintPreviewDialog { Document = pd, Width = 1024, Height = 768, WindowState = FormWindowState.Maximized };
            ppd.ShowDialog();
        }
    }
}
