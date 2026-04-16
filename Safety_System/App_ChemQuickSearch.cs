/// FILE: Safety_System/App_ChemQuickSearch.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        // 成員控制項
        private DataGridView _dgvResult;
        private TextBox _txtName;
        private TextBox _txtCAS;
        
        // 常數定義
        private const string DbName = "Chemical";
        private const string TableName = "ChemRegulations";

        public Control GetView()
        {
            // 主容器：使用 TableLayoutPanel 進行縱向佈局
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(20), 
                RowCount = 2,
                BackColor = Color.WhiteSmoke
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 第一行：功能按鈕
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 第二行：主大框

            // --- 第一行：功能按鈕區 ---
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 15) 
            };
            
            Button btnPdf = new Button { 
                Text = "📄 導出快查分析 PDF", 
                Size = new Size(200, 45), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
            btnPdf.Click += (s, e) => ExportToPdf();
            pnlAction.Controls.Add(btnPdf);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // --- 第二行：第一大框 (包含 4 個小框) ---
            GroupBox boxMain = new GroupBox { 
                Text = "🔍 化學品法規快查中心", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.DarkCyan, 
                Padding = new Padding(15) 
            };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 65F));  // 小框1: 標題
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 85F));  // 小框2: 名稱查詢
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 85F));  // 小框3: CAS查詢
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 小框4: 結果顯示

            // 【小框 1】：標題文字
            Panel sub1 = CreateSubBox("化學品快查分析", Color.Teal);
            Label lblMainTitle = new Label { 
                Text = "🧬 化學品法規要求與成分快查分析系統", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                ForeColor = Color.Teal 
            };
            sub1.Controls.Add(lblMainTitle);

            // 【小框 2】：化學品名稱查詢
            Panel sub2 = CreateSubBox("按名稱檢索", Color.FromArgb(45, 62, 80));
            Label lbl1 = new Label { Text = "化學品名稱關鍵字：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtName = new TextBox { Location = new Point(200, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtName.TextChanged += (s, e) => ExecuteSearch(); // 即時搜尋
            sub2.Controls.AddRange(new Control[] { lbl1, _txtName });

            // 【小框 3】：CAS No 查詢
            Panel sub3 = CreateSubBox("按 CAS No 檢索", Color.FromArgb(45, 62, 80));
            Label lbl2 = new Label { Text = "CAS No. 編號：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtCAS = new TextBox { Location = new Point(200, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtCAS.TextChanged += (s, e) => ExecuteSearch(); // 即時搜尋
            sub3.Controls.AddRange(new Control[] { lbl2, _txtCAS });

            // 【小框 4】：查詢結果顯示
            GroupBox sub4 = new GroupBox { 
                Text = "📊 檢索結果明細 (自動過濾空白欄位)", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                Margin = new Padding(0, 10, 0, 0) 
            };
            _dgvResult = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None
            };
            _dgvResult.RowTemplate.Height = 35;
            _dgvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            _dgvResult.EnableHeadersVisualStyles = false;
            _dgvResult.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
            _dgvResult.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvResult.ColumnHeadersHeight = 40;
            sub4.Controls.Add(_dgvResult);

            // 組裝小框
            innerTable.Controls.Add(sub1, 0, 0);
            innerTable.Controls.Add(sub2, 0, 1);
            innerTable.Controls.Add(sub3, 0, 2);
            innerTable.Controls.Add(sub4, 0, 3);
            
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            // 初始執行一次空搜尋，確保載入結構
            ExecuteSearch(); 

            return mainLayout;
        }

        /// <summary>
        /// 建立美化的小框容器
        /// </summary>
        private Panel CreateSubBox(string title, Color accentColor)
        {
            Panel p = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), BackColor = Color.White };
            p.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.FromArgb(220, 220, 220), ButtonBorderStyle.Solid);
                using (SolidBrush brush = new SolidBrush(accentColor)) {
                    e.Graphics.FillRectangle(brush, 0, 0, 6, p.Height); // 左側裝飾條
                }
            };
            return p;
        }

        /// <summary>
        /// 執行檢索邏輯：包含動態欄位隱藏功能
        /// </summary>
        private void ExecuteSearch()
        {
            try
            {
                string nameKey = _txtName.Text.Trim();
                string casKey = _txtCAS.Text.Trim();

                // 從資料庫抓取全表
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                if (dt == null || dt.Columns.Count == 0) {
                    _dgvResult.DataSource = null;
                    return;
                }

                DataView dv = dt.DefaultView;
                List<string> filters = new List<string>();
                
                // 支援包含文字查詢
                if (!string.IsNullOrEmpty(nameKey)) 
                    filters.Add($"([中文名稱] LIKE '%{nameKey}%' OR [英文名稱] LIKE '%{nameKey}%' OR [法規名稱] LIKE '%{nameKey}%')");
                
                if (!string.IsNullOrEmpty(casKey)) 
                    filters.Add($"([CAS No] LIKE '%{casKey}%')");

                if (filters.Count > 0) 
                    dv.RowFilter = string.Join(" AND ", filters);
                
                DataTable resultTable = dv.ToTable();
                _dgvResult.DataSource = resultTable;
                
                // 固定隱藏 Id
                if (_dgvResult.Columns.Contains("Id")) _dgvResult.Columns["Id"].Visible = false;

                // 🟢 動態欄位隱藏：檢查每一欄是否有資料
                foreach (DataGridViewColumn col in _dgvResult.Columns)
                {
                    if (col.Name == "Id") continue;

                    bool hasData = false;
                    foreach (DataRow row in resultTable.Rows)
                    {
                        if (row[col.Name] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col.Name].ToString()))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    col.Visible = hasData; // 僅顯示有資料的欄位
                }
            }
            catch { _dgvResult.DataSource = null; }
        }

        /// <summary>
        /// 分頁導出 PDF
        /// </summary>
        private void ExportToPdf()
        {
            if (_dgvResult.DataSource == null || _dgvResult.Rows.Count == 0) {
                MessageBox.Show("目前暫無搜尋數據可供導出報表。");
                return;
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; // 橫向
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int rowIndex = 0;
            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                // 1. 報表頁首
                g.DrawString("化學品法規快查分析報表", fTitle, Brushes.DarkSlateGray, x, y);
                y += 45;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   台灣玻璃工業股份有限公司-彰濱廠   |   頁碼：{rowIndex / 25 + 1}", fBody, Brushes.Gray, x, y);
                y += 35;

                // 2. 準備列印欄位 (僅列印目前看板上「可見」的欄位)
                var visCols = _dgvResult.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                if (visCols.Count == 0) return;

                // 計算縮放比例
                float totalW = visCols.Sum(c => c.Width);
                float scale = e.MarginBounds.Width / totalW;
                if (scale > 1.2f) scale = 1.2f;

                // 3. 繪製表頭
                float currX = x;
                float rowH = 30;
                foreach (var col in visCols) {
                    RectangleF rect = new RectangleF(currX, y, col.Width * scale, rowH);
                    g.FillRectangle(Brushes.LightGray, rect);
                    g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                    g.DrawString(col.HeaderText, fHead, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    currX += col.Width * scale;
                }
                y += rowH;

                // 4. 繪製內容 (逐列列印並檢查分頁)
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

                    // 檢查分頁高度
                    if (y + rowH > e.MarginBounds.Bottom) {
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
                WindowState = FormWindowState.Maximized 
            };
            ppd.ShowDialog();
        }
    }
}
