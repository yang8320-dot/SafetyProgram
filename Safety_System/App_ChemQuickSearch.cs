using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        private TextBox _txtName;
        private TextBox _txtCAS;
        private Button _btnSearch;
        private Label _lblStatus;
        private FlowLayoutPanel _flpResultsContainer; 
        
        private const string DbName = "Chemical";

        // 內部類別用於定義 12 個表的查詢與顯示規則
        private class ChemTableInfo {
            public string TableName;
            public string Title;
            public string NameSearchCol;
            public string CasSearchCol;
            public string ExtraNotice;
            public GroupBox GBox;
            public DataGridView Dgv;
            public DataTable ResultData;
        }

        private List<ChemTableInfo> _tableInfos = new List<ChemTableInfo> {
            new ChemTableInfo { TableName="EnvTesting", Title="1. 環測項目", NameSearchCol="中文名稱", CasSearchCol="CASNO", ExtraNotice="" },
            new ChemTableInfo { TableName="ExposureLimits", Title="2. 勞工暴露容許濃度", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="ToxicSubstances", Title="3. 毒性物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="ConcernedChem", Title="4. 關注性化學物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="PriorityMgmtChem", Title="5. 優先管理化學品", NameSearchCol="中文名稱", CasSearchCol="CASNO", ExtraNotice="" },
            new ChemTableInfo { TableName="ControlledChem", Title="6. 管制化學品", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="SpecificChem", Title="7. 特定化學物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="需設置【特化主管】" },
            new ChemTableInfo { TableName="OrganicSolvents", Title="8. 有機溶劑", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="需設置【有機溶劑作業主管】" },
            new ChemTableInfo { TableName="WorkerHealthProtect", Title="9. 勞工健康保護", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="需【特殊體檢】" },
            new ChemTableInfo { TableName="PublicHazardous", Title="10. 公共危險物品", NameSearchCol="種類", CasSearchCol="種類", ExtraNotice="" },
            new ChemTableInfo { TableName="AirPollutionEmerg", Title="11. 空污緊急應變", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="FactoryHazardous", Title="12. 工廠危險物品申報", NameSearchCol="種類", CasSearchCol="種類", ExtraNotice="" }
        };

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(20), 
                RowCount = 2,
                BackColor = Color.WhiteSmoke
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // --- 第一行：功能按鈕區 ---
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 15) 
            };
            
            Button btnPdf = new Button { 
                Text = "📄 導出分析 PDF", 
                Size = new Size(180, 45), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
            btnPdf.Click += (s, e) => ExportToPdf();

            _btnSearch = new Button {
                Text = "🚀 開始執行交叉檢索",
                Size = new Size(220, 45),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(15, 0, 0, 0)
            };
            // 綁定非同步查詢事件
            _btnSearch.Click += async (s, e) => await ExecuteSearchAsync();

            _lblStatus = new Label {
                Text = "準備就緒。請輸入條件後點擊查詢。",
                ForeColor = Color.DimGray,
                Font = new Font("Microsoft JhengHei UI", 11F),
                AutoSize = true,
                Margin = new Padding(15, 12, 0, 0)
            };

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(_btnSearch);
            pnlAction.Controls.Add(_lblStatus);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // --- 第二行：主查詢區與結果區 ---
            GroupBox boxMain = new GroupBox { 
                Text = "🔍 化學品法規快查中心", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.DarkCyan, 
                Padding = new Padding(15) 
            };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 65F));  
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 85F));  
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 85F));  
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // 【小框 1】：標題文字
            Panel sub1 = CreateSubBox(Color.Teal);
            Label lblMainTitle = new Label { 
                Text = "🧬 化學品法規符核度查詢系統", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), 
                ForeColor = Color.Teal 
            };
            sub1.Controls.Add(lblMainTitle);

            // 【小框 2】：名稱查詢
            Panel sub2 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl1 = new Label { Text = "化學品名稱關鍵字：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtName = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 14F) };
            // 已移除 TextChanged 連動，改由按鈕觸發
            sub2.Controls.AddRange(new Control[] { lbl1, _txtName });

            // 【小框 3】：CAS 查詢
            Panel sub3 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl2 = new Label { Text = "CAS No. 編號：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtCAS = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 14F) };
            // 已移除 TextChanged 連動，改由按鈕觸發
            sub3.Controls.AddRange(new Control[] { lbl2, _txtCAS });

            // 【小框 4】：檢索結果容器
            GroupBox sub4 = new GroupBox { 
                Text = "📊 檢索結果明細 (查無資料之分類將自動隱藏)", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                Margin = new Padding(0, 10, 0, 0) 
            };

            _flpResultsContainer = new FlowLayoutPanel {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10),
                BackColor = Color.White
            };
            
            // 🟢 初始化 12 個資料表的視窗結構 (修正高度計算坍塌 Bug)
            foreach (var info in _tableInfos) {
                info.GBox = new GroupBox {
                    Text = info.Title + (string.IsNullOrEmpty(info.ExtraNotice) ? "" : " - " + info.ExtraNotice),
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                    ForeColor = string.IsNullOrEmpty(info.ExtraNotice) ? Color.DarkSlateBlue : Color.Crimson,
                    Width = 1100,
                    AutoSize = false, // 🛑 關閉自動縮放，改由程式手動計算高度避免坍塌
                    Margin = new Padding(0, 0, 0, 25),
                    Padding = new Padding(10, 30, 10, 10), // 留出標題空間
                    Visible = false 
                };

                info.Dgv = new DataGridView { 
                    Dock = DockStyle.Fill, // 🛑 確保表格填滿手動設定高度的 GroupBox
                    BackgroundColor = Color.White, 
                    AllowUserToAddRows = false, 
                    ReadOnly = true, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                    RowHeadersVisible = false,
                    BorderStyle = BorderStyle.FixedSingle,
                    ScrollBars = ScrollBars.Both, // 允許出現捲軸
                    Font = new Font("Microsoft JhengHei UI", 11F)
                };
                info.Dgv.RowTemplate.Height = 35;
                info.Dgv.EnableHeadersVisualStyles = false;
                info.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
                info.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                info.Dgv.ColumnHeadersHeight = 40;
                info.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;

                info.GBox.Controls.Add(info.Dgv);
                _flpResultsContainer.Controls.Add(info.GBox);
            }

            sub4.Controls.Add(_flpResultsContainer);
            innerTable.Controls.Add(sub1, 0, 0);
            innerTable.Controls.Add(sub2, 0, 1);
            innerTable.Controls.Add(sub3, 0, 2);
            innerTable.Controls.Add(sub4, 0, 3);
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            return mainLayout;
        }

        private Panel CreateSubBox(Color accentColor)
        {
            Panel p = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), BackColor = Color.Transparent };
            p.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.FromArgb(200, 200, 200), ButtonBorderStyle.Solid);
                using (SolidBrush brush = new SolidBrush(accentColor)) {
                    e.Graphics.FillRectangle(brush, 0, 0, 6, p.Height); 
                }
            };
            return p;
        }

        private async Task ExecuteSearchAsync()
        {
            string nameKey = _txtName.Text.Trim();
            string casKey = _txtCAS.Text.Trim();

            if (string.IsNullOrEmpty(nameKey) && string.IsNullOrEmpty(casKey)) {
                MessageBox.Show("請至少輸入一個查詢關鍵字或 CAS No。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 鎖定按鈕，防止重複點擊
            _btnSearch.Enabled = false;
            _btnSearch.Text = "⏳ 檢索中...";
            _lblStatus.Text = "正在背景檢索 12 個資料庫表，請稍候...";
            _lblStatus.ForeColor = Color.OrangeRed;

            // 🟢 使用 Task.Run 進行背景非同步運算，防止 UI 執行緒假當機
            await Task.Run(() => {
                foreach (var info in _tableInfos) {
                    try {
                        DataTable dt = DataManager.GetTableData(DbName, info.TableName, "", "", "");
                        if (dt != null && dt.Rows.Count > 0) {
                            DataView dv = dt.DefaultView;
                            List<string> filters = new List<string>();
                            
                            if (!string.IsNullOrEmpty(nameKey) && dt.Columns.Contains(info.NameSearchCol)) 
                                filters.Add($"[{info.NameSearchCol}] LIKE '%{nameKey.Replace("'", "''")}%'");
                            
                            if (!string.IsNullOrEmpty(casKey) && dt.Columns.Contains(info.CasSearchCol)) 
                                filters.Add($"[{info.CasSearchCol}] LIKE '%{casKey.Replace("'", "''")}%'");

                            if (filters.Count > 0) 
                                dv.RowFilter = string.Join(" AND ", filters);
                            else
                                dv.RowFilter = "1=0"; // 若沒輸條件則該表不顯示
                            
                            info.ResultData = dv.ToTable();
                        } else {
                            info.ResultData = null;
                        }
                    } catch {
                        info.ResultData = null;
                    }
                }
            });

            // 回到 UI 執行緒更新畫面
            _flpResultsContainer.SuspendLayout();
            int totalFound = 0;

            foreach (var info in _tableInfos) {
                if (info.ResultData != null && info.ResultData.Rows.Count > 0) {
                    
                    info.GBox.Visible = true; // 先設為可見，確保 DataGridView 計算排版時能抓到實體屬性
                    info.Dgv.DataSource = info.ResultData;

                    if (info.Dgv.Columns.Contains("Id")) info.Dgv.Columns["Id"].Visible = false;

                    // 動態隱藏空白欄位
                    foreach (DataGridViewColumn col in info.Dgv.Columns) {
                        if (col.Name == "Id") continue;
                        bool hasValue = false;
                        foreach (DataRow row in info.ResultData.Rows) {
                            if (row[col.Name] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col.Name].ToString())) {
                                hasValue = true; break;
                            }
                        }
                        col.Visible = hasValue;
                    }

                    // 🛑 強制手動計算並設定 GroupBox 的精準高度，徹底解決表格空白/坍塌的問題
                    int rowCount = info.ResultData.Rows.Count;
                    int targetHeight = info.Dgv.ColumnHeadersHeight + (rowCount * info.Dgv.RowTemplate.Height) + info.GBox.Padding.Top + info.GBox.Padding.Bottom + 10;
                    
                    // 最高限制 400px，超過的由 ScrollBars.Both 接手
                    info.GBox.Height = Math.Min(targetHeight, 400); 
                    
                    totalFound += rowCount;
                } else {
                    info.GBox.Visible = false;
                    info.Dgv.DataSource = null; // 清除舊資料
                }
            }

            _flpResultsContainer.ResumeLayout(true);
            _btnSearch.Enabled = true;
            _btnSearch.Text = "🚀 開始執行交叉檢索";
            _lblStatus.Text = $"檢索完成！共在各分類中找到 {totalFound} 筆資料。";
            _lblStatus.ForeColor = Color.ForestGreen;

            if (totalFound == 0) {
                MessageBox.Show("於所有法規資料庫中皆查無符合條件之項目。", "檢索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ExportToPdf()
        {
            var visibleTables = _tableInfos.Where(t => t.GBox.Visible && t.ResultData != null && t.ResultData.Rows.Count > 0).ToList();
            if (visibleTables.Count == 0) {
                MessageBox.Show("目前暫無搜尋數據可供導出。"); return;
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; 
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int currentTableIndex = 0;
            int currentRowIndex = 0;

            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fSubTitle = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                g.DrawString("化學品法規符核度分析報表", fTitle, Brushes.Black, x, y);
                y += 45;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   台灣玻璃彰濱廠", fBody, Brushes.Gray, x, y);
                y += 35;

                while (currentTableIndex < visibleTables.Count) {
                    var info = visibleTables[currentTableIndex];
                    var visCols = info.Dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    
                    if (currentRowIndex == 0) {
                        if (y + 100 > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                        g.DrawString(info.Title + " " + info.ExtraNotice, fSubTitle, Brushes.DarkSlateBlue, x, y);
                        y += 30;

                        float currX = x;
                        float totalW = visCols.Sum(c => c.Width);
                        float scale = Math.Min(1.2f, e.MarginBounds.Width / totalW);

                        foreach (var col in visCols) {
                            RectangleF rect = new RectangleF(currX, y, col.Width * scale, 30);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            g.DrawString(col.HeaderText, fHead, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                            currX += col.Width * scale;
                        }
                        y += 30;
                    }

                    float tW = visCols.Sum(c => c.Width);
                    float tS = Math.Min(1.2f, e.MarginBounds.Width / tW);

                    while (currentRowIndex < info.ResultData.Rows.Count) {
                        if (y + 30 > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                        float currX = x;
                        foreach (var col in visCols) {
                            RectangleF rect = new RectangleF(currX, y, col.Width * tS, 30);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            string val = info.Dgv[col.Index, currentRowIndex].Value?.ToString() ?? "";
                            g.DrawString(val, fBody, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
                            currX += col.Width * tS;
                        }
                        y += 30;
                        currentRowIndex++;
                    }
                    y += 20; currentTableIndex++; currentRowIndex = 0;
                }
                e.HasMorePages = false; currentTableIndex = 0; currentRowIndex = 0;
            };

            PrintPreviewDialog ppd = new PrintPreviewDialog { Document = pd, Width = 1024, Height = 768, WindowState = FormWindowState.Maximized };
            ppd.ShowDialog();
        }
    }
}
