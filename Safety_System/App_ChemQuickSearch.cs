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

        private class ChemTableInfo {
            public string TableName;
            public string Title;
            public string NameSearchCol;
            public string CasSearchCol;
            public string ExtraNotice;
            public GroupBox GBox;
            public DataGridView Dgv;
            public DataTable ResultData;
            public List<string> VisibleColumns; 
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

            // 🟢 優化：水平對齊功能列的按鈕與文字
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 10),
                WrapContents = false 
            };
            
            Button btnPdf = new Button { 
                Text = "📄 導出分析 PDF", 
                Size = new Size(180, 40), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 5, 10, 5) // 統一上下邊界
            };
            btnPdf.Click += (s, e) => ExportToPdf();

            _btnSearch = new Button {
                Text = "🚀 開始執行交叉檢索",
                Size = new Size(220, 40),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5, 5, 15, 5) // 統一上下邊界
            };
            _btnSearch.Click += async (s, e) => await ExecuteSearchAsync();

            _lblStatus = new Label {
                Text = "準備就緒。請輸入條件後點擊查詢或按下 Enter 鍵。",
                ForeColor = Color.DimGray,
                Font = new Font("Microsoft JhengHei UI", 11F),
                AutoSize = true,
                Margin = new Padding(0, 15, 0, 0) // 下沉對齊按鈕文字的基準線
            };

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(_btnSearch);
            pnlAction.Controls.Add(_lblStatus);
            mainLayout.Controls.Add(pnlAction, 0, 0);

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

            Panel sub1 = CreateSubBox(Color.Teal);
            Label lblMainTitle = new Label { 
                Text = "🧬 化學品法規符核度查詢系統", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), 
                ForeColor = Color.Teal 
            };
            sub1.Controls.Add(lblMainTitle);

            Panel sub2 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl1 = new Label { Text = "化學品名稱關鍵字：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtName = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 14F) };
            // 🟢 新增 Enter 鍵觸發查詢
            _txtName.KeyDown += async (s, e) => {
                if (e.KeyCode == Keys.Enter) {
                    e.Handled = true; e.SuppressKeyPress = true; // 消除系統提示音
                    await ExecuteSearchAsync();
                }
            };
            sub2.Controls.AddRange(new Control[] { lbl1, _txtName });

            Panel sub3 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl2 = new Label { Text = "CAS No. 編號：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _txtCAS = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 14F) };
            // 🟢 新增 Enter 鍵觸發查詢
            _txtCAS.KeyDown += async (s, e) => {
                if (e.KeyCode == Keys.Enter) {
                    e.Handled = true; e.SuppressKeyPress = true;
                    await ExecuteSearchAsync();
                }
            };
            sub3.Controls.AddRange(new Control[] { lbl2, _txtCAS });

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
                Padding = new Padding(10, 10, 30, 10), 
                BackColor = Color.White
            };
            
            _flpResultsContainer.Resize += (s, e) => {
                foreach (Control c in _flpResultsContainer.Controls) {
                    if (c is GroupBox gb) gb.Width = _flpResultsContainer.ClientSize.Width - 40;
                }
            };
            
            foreach (var info in _tableInfos) {
                info.GBox = new GroupBox {
                    Text = info.Title + (string.IsNullOrEmpty(info.ExtraNotice) ? "" : " - " + info.ExtraNotice),
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                    ForeColor = string.IsNullOrEmpty(info.ExtraNotice) ? Color.DarkSlateBlue : Color.Crimson,
                    AutoSize = false, 
                    Margin = new Padding(0, 0, 0, 20),
                    Padding = new Padding(8, 35, 8, 8), 
                    Visible = false 
                };

                info.Dgv = new DataGridView { 
                    Dock = DockStyle.Fill, 
                    BackgroundColor = Color.White, 
                    AllowUserToAddRows = false, 
                    ReadOnly = true, 
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                    RowHeadersVisible = false,
                    BorderStyle = BorderStyle.FixedSingle,
                    ScrollBars = ScrollBars.None, 
                    Font = new Font("Microsoft JhengHei UI", 11F)
                };
                
                // 🟢 解決顏色不均：統一底色與反白顏色設計
                info.Dgv.RowTemplate.Height = 35; 
                info.Dgv.EnableHeadersVisualStyles = false;
                info.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
                info.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                info.Dgv.ColumnHeadersHeight = 40; 
                
                info.Dgv.DefaultCellStyle.BackColor = Color.White;
                info.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                // 設定柔和的反白顏色，並確保所有欄位顏色一致
                info.Dgv.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue; 
                info.Dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

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

            _btnSearch.Enabled = false;
            _btnSearch.Text = "⏳ 檢索與排版中...";
            _lblStatus.Text = "正在背景非同步檢索資料庫，請稍候...";
            _lblStatus.ForeColor = Color.OrangeRed;

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

                            dv.RowFilter = filters.Count > 0 ? string.Join(" AND ", filters) : "1=0";
                            info.ResultData = dv.ToTable();

                            info.VisibleColumns = new List<string>();
                            if (info.ResultData.Rows.Count > 0) {
                                foreach (DataColumn col in info.ResultData.Columns) {
                                    if (col.ColumnName == "Id") continue;
                                    bool hasValue = false;
                                    foreach (DataRow row in info.ResultData.Rows) {
                                        if (row[col] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col].ToString())) {
                                            hasValue = true; break;
                                        }
                                    }
                                    if (hasValue) info.VisibleColumns.Add(col.ColumnName);
                                }
                            }
                        } else {
                            info.ResultData = null;
                        }
                    } catch {
                        info.ResultData = null;
                    }
                }
            });

            _flpResultsContainer.SuspendLayout();
            int totalFound = 0;

            foreach (var info in _tableInfos) {
                if (info.ResultData != null && info.ResultData.Rows.Count > 0) {
                    
                    info.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                    info.Dgv.DataSource = info.ResultData;

                    foreach (DataGridViewColumn col in info.Dgv.Columns) {
                        col.Visible = info.VisibleColumns.Contains(col.Name);
                    }

                    info.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

                    int rowCount = info.ResultData.Rows.Count;
                    int exactGridHeight = info.Dgv.ColumnHeadersHeight + (rowCount * info.Dgv.RowTemplate.Height);
                    int targetGBoxHeight = 35 + exactGridHeight + 8 + 20; 

                    info.GBox.Width = _flpResultsContainer.ClientSize.Width - 40;
                    info.GBox.Height = targetGBoxHeight; 
                    
                    // 🟢 解決顏色不一致：清除預設的反白選取，讓表格顏色回歸自然
                    info.Dgv.ClearSelection();
                    
                    info.GBox.Visible = true;
                    totalFound += rowCount;

                    await Task.Delay(10); 
                } else {
                    info.GBox.Visible = false;
                    info.Dgv.DataSource = null; 
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
