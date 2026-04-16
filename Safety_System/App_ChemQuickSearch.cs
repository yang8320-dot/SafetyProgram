/// FILE: Safety_System/App_ChemQuickSearch.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing; // 🟢 補回此關鍵引用
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        private DataGridView _dgvResult;
        private TextBox _txtName, _txtCAS;
        private const string DbName = "Chemical";
        private const string TableName = "ChemRegulations";

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 第一行：功能按鈕區
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            Button btnPdf = new Button { Text = "📄 導出 PDF 報表", Size = new Size(180, 40), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.Click += (s, e) => ExportToPdf();
            pnlAction.Controls.Add(btnPdf);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // 第二行：主大框 (內含4個小框)
            GroupBox boxMain = new GroupBox { Text = "🔍 化學品快速檢索中心", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkCyan, Padding = new Padding(15) };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));  // 小框1: 標題
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));  // 小框2: 名稱查詢
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));  // 小框3: CAS查詢
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 小框4: 結果顯示

            // 小框 1：標題文字
            Panel sub1 = CreateSubBox("化學品快查分析", Color.Teal);
            Label lblMainTitle = new Label { Text = "🧬 化學品法規要求與成分快查分析", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.Teal };
            sub1.Controls.Add(lblMainTitle);

            // 小框 2：化學品名稱查詢
            Panel sub2 = CreateSubBox("按化學品名稱查詢", Color.FromArgb(45, 62, 80));
            Label lbl1 = new Label { Text = "請輸入名稱關鍵字：", Location = new Point(20, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtName = new TextBox { Location = new Point(180, 17), Width = 400, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtName.TextChanged += (s, e) => ExecuteSearch();
            sub2.Controls.AddRange(new Control[] { lbl1, _txtName });

            // 小框 3：CAS No 查詢
            Panel sub3 = CreateSubBox("按 CAS No. 查詢", Color.FromArgb(45, 62, 80));
            Label lbl2 = new Label { Text = "請輸入 CAS 編號：", Location = new Point(20, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtCAS = new TextBox { Location = new Point(180, 17), Width = 400, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtCAS.TextChanged += (s, e) => ExecuteSearch();
            sub3.Controls.AddRange(new Control[] { lbl2, _txtCAS });

            // 小框 4：搜尋結果與數據網格
            GroupBox sub4 = new GroupBox { Text = "📊 檢索結果數據明細 (僅顯示有資料之欄位)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Margin = new Padding(0, 10, 0, 0) };
            _dgvResult = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false };
            _dgvResult.RowTemplate.Height = 35;
            _dgvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            sub4.Controls.Add(_dgvResult);

            innerTable.Controls.Add(sub1, 0, 0);
            innerTable.Controls.Add(sub2, 0, 1);
            innerTable.Controls.Add(sub3, 0, 2);
            innerTable.Controls.Add(sub4, 0, 3);
            
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            ExecuteSearch(); 
            return mainLayout;
        }

        private Panel CreateSubBox(string title, Color borderColor)
        {
            Panel p = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), BackColor = Color.White };
            p.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
                e.Graphics.FillRectangle(new SolidBrush(borderColor), 0, 0, 5, p.Height); 
            };
            return p;
        }

        private void ExecuteSearch()
        {
            string nameKey = _txtName.Text.Trim();
            string casKey = _txtCAS.Text.Trim();

            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            if (dt == null) return;
            DataView dv = dt.DefaultView;
            
            List<string> filters = new List<string>();
            if (!string.IsNullOrEmpty(nameKey)) filters.Add($"([中文名稱] LIKE '%{nameKey}%' OR [英文名稱] LIKE '%{nameKey}%' OR [法規名稱] LIKE '%{nameKey}%')");
            if (!string.IsNullOrEmpty(casKey)) filters.Add($"([CAS No] LIKE '%{casKey}%')");

            if (filters.Count > 0) dv.RowFilter = string.Join(" AND ", filters);
            
            DataTable filteredDt = dv.ToTable();

            _dgvResult.DataSource = filteredDt;
            
            if (_dgvResult.Columns.Contains("Id")) _dgvResult.Columns["Id"].Visible = false;

            foreach (DataGridViewColumn col in _dgvResult.Columns)
            {
                bool hasData = false;
                foreach (DataRow row in filteredDt.Rows)
                {
                    if (row[col.Name] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col.Name].ToString()))
                    {
                        hasData = true;
                        break;
                    }
                }
                col.Visible = hasData;
            }
        }

        private void ExportToPdf()
        {
            if (_dgvResult.Rows.Count == 0) { MessageBox.Show("無搜尋結果可匯出。"); return; }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true;
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int rowIndex = 0;
            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                g.DrawString("化學品法規快查分析報表", fTitle, Brushes.Black, x, y);
                y += 40;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm} | 台灣玻璃工業股份有限公司-彰濱廠", fBody, Brushes.Gray, x, y);
                y += 30;

                var visCols = _dgvResult.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                float totalW = visCols.Sum(c => c.Width);
                float scale = e.MarginBounds.Width / totalW;
                if (scale > 1) scale = 1;

                float currX = x;
                float rowH = 30;
                foreach (var col in visCols) {
                    RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                    g.FillRectangle(Brushes.LightSteelBlue, rect);
                    g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                    g.DrawString(col.HeaderText, fHead, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    currX += col.Width * scale;
                }
                y += rowH;

                while (rowIndex < _dgvResult.Rows.Count) {
                    currX = x;
                    foreach (var col in visCols) {
                        RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                        g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                        string val = _dgvResult[col.Index, rowIndex].Value?.ToString() ?? "";
                        g.DrawString(val, fBody, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
                        currX += col.Width * scale;
                    }
                    y += rowH;
                    rowIndex++;

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
