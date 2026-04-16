/// FILE: Safety_System/App_ChemDashboard.cs ///
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemDashboard
    {
        private DataGridView _dgvSDS;
        private const string DbName = "Chemical";
        private const string TableName = "SDS_Inventory";

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 第一行：功能按鈕
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            Button btnPdf = new Button { Text = "📤 導出全廠 SDS 清冊 PDF", Size = new Size(240, 40), BackColor = Color.DarkCyan, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.Click += (s, e) => ExportToPdf();
            pnlAction.Controls.Add(btnPdf);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // 第二大框
            GroupBox boxMain = new GroupBox { Text = "📋 化學品管理綜合看板", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Padding = new Padding(15) };
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F)); // 小框1: 標題區
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 小框2: 數據區

            // 小框 1：標題文字
            Panel pnlTitle = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 245, 250), Margin = new Padding(0,0,0,15) };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Solid);
            
            Label lblCompany = new Label { Text = "台灣玻璃工業股份有限公司 - 彰濱廠", Dock = DockStyle.Top, Height = 45, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.MidnightBlue };
            Label lblSubTitle = new Label { Text = "化學品清單一覽表 (SDS 定期追蹤管理)", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.FromArgb(60, 60, 60) };
            pnlTitle.Controls.Add(lblSubTitle);
            pnlTitle.Controls.Add(lblCompany);

            // 小框 2：SDS 數據表格
            GroupBox boxGrid = new GroupBox { Text = "SDS 安全資料表庫存清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            _dgvSDS = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, BorderStyle = BorderStyle.None };
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
            if (_dgvSDS.Columns.Contains("Id")) _dgvSDS.Columns["Id"].Visible = false;
            if (_dgvSDS.Columns.Contains("附件檔案")) {
                _dgvSDS.Columns["附件檔案"].Visible = false; // 看板不顯示檔案路徑
            }
        }

        private void ExportToPdf()
        {
            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true;
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
                        g.DrawString(_dgvSDS[col.Index, rowIndex].Value?.ToString() ?? "", new Font("Microsoft JhengHei UI", 9F), Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
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
