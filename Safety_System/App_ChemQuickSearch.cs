/// FILE: Safety_System/App_ChemQuickSearch.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        private TextBox _txtName;
        private TextBox _txtCAS;
        
        // 將 12 個資料表結果統一動態生成並放置於此面板
        private FlowLayoutPanel _flpResultsContainer; 
        
        private const string DbName = "Chemical";

        // 定義 12 個表的查詢規則
        private class ChemTableInfo {
            public string TableName;
            public string Title;
            public string NameSearchCol;
            public string CasSearchCol;
            public string ExtraNotice;
            public GroupBox GBox;
            public DataGridView Dgv;
        }

        // 🟢 嚴格依照需求：配置 12 個表的欄位對應與附註文字
        private List<ChemTableInfo> _tableInfos = new List<ChemTableInfo> {
            new ChemTableInfo { TableName="EnvTesting", Title="1. 環測項目", NameSearchCol="中文名稱", CasSearchCol="CASNO", ExtraNotice="" },
            new ChemTableInfo { TableName="ExposureLimits", Title="2. 勞工暴露容許濃度", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="ToxicSubstances", Title="3. 毒性物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="ConcernedChem", Title="4. 關注性化學物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="PriorityMgmtChem", Title="5. 優先管理化學品", NameSearchCol="中文名稱", CasSearchCol="CASNO", ExtraNotice="" },
            new ChemTableInfo { TableName="ControlledChem", Title="6. 管制化學品", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice="" },
            new ChemTableInfo { TableName="SpecificChem", Title="7. 特定化學物質", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice=" (需設置【特化主管】)" },
            new ChemTableInfo { TableName="OrganicSolvents", Title="8. 有機溶劑", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice=" (需設置【有機溶劑作業主管】)" },
            new ChemTableInfo { TableName="WorkerHealthProtect", Title="9. 勞工健康保護", NameSearchCol="中文名稱", CasSearchCol="中文名稱", ExtraNotice=" (需【特殊體檢】)" },
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

            // --- 第二行：主大框 ---
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
                Text = "🧬 化學品法規符核度查詢系統", // 🟢 依需求變更名稱
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                ForeColor = Color.Teal 
            };
            sub1.Controls.Add(lblMainTitle);

            // 【小框 2】：名稱查詢
            Panel sub2 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl1 = new Label { Text = "化學品名稱關鍵字：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            // 🟢 X座標由 200 改為 220，避免重疊並增加間隔
            _txtName = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtName.TextChanged += (s, e) => ExecuteSearch(); 
            sub2.Controls.AddRange(new Control[] { lbl1, _txtName });

            // 【小框 3】：CAS 查詢
            Panel sub3 = CreateSubBox(Color.FromArgb(45, 62, 80));
            Label lbl2 = new Label { Text = "CAS No. 編號：", Location = new Point(25, 25), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            // 🟢 X座標由 200 改為 220，對齊上方
            _txtCAS = new TextBox { Location = new Point(220, 22), Width = 450, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtCAS.TextChanged += (s, e) => ExecuteSearch(); 
            sub3.Controls.AddRange(new Control[] { lbl2, _txtCAS });

            // 【小框 4】：檢索結果容器 (放置 12 個小框)
            GroupBox sub4 = new GroupBox { 
                Text = "📊 檢索結果明細 (查無資料之法規表單將自動隱藏)", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                Margin = new Padding(0, 10, 0, 0) 
            };

            _flpResultsContainer = new FlowLayoutPanel {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10)
            };
            
            // 動態產生 12 個表的 UI 容器
            foreach (var info in _tableInfos) {
                info.GBox = new GroupBox {
                    Text = info.Title + info.ExtraNotice,
                    Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                    ForeColor = Color.DarkSlateBlue,
                    Width = 1200,
                    AutoSize = true,
                    Margin = new Padding(0, 0, 0, 20),
                    Padding = new Padding(5, 10, 5, 5),
                    Visible = false // 預設隱藏
                };

                info.Dgv = new DataGridView { 
                    Dock = DockStyle.Fill, 
                    MinimumSize = new Size(1180, 80),
                    BackgroundColor = Color.White, 
                    AllowUserToAddRows = false, 
                    ReadOnly = true, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                    RowHeadersVisible = false,
                    BorderStyle = BorderStyle.None
                };
                info.Dgv.RowTemplate.Height = 30;
                info.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                info.Dgv.EnableHeadersVisualStyles = false;
                info.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
                info.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                info.Dgv.ColumnHeadersHeight = 35;

                info.Dgv.DataBindingComplete += (s, e) => {
                    // 自動調整高度以展開所有資料 (最高限制避免過長)
                    int height = info.Dgv.ColumnHeadersHeight;
                    foreach (DataGridViewRow row in info.Dgv.Rows) height += row.Height;
                    info.Dgv.Height = Math.Min(height + 10, 300);
                };

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

            ExecuteSearch(); 

            return mainLayout;
        }

        // 🟢 取消白底：改為 Color.Transparent
        private Panel CreateSubBox(Color accentColor)
        {
            Panel p = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), BackColor = Color.Transparent };
            p.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.FromArgb(220, 220, 220), ButtonBorderStyle.Solid);
                using (SolidBrush brush = new SolidBrush(accentColor)) {
                    e.Graphics.FillRectangle(brush, 0, 0, 6, p.Height); 
                }
            };
            return p;
        }

        private void ExecuteSearch()
        {
            string nameKey = _txtName.Text.Trim();
            string casKey = _txtCAS.Text.Trim();

            // 🟢 優化：暫停 UI 繪製，防止輸入文字時因頻繁刷新畫面導致卡頓或無反應
            _flpResultsContainer.SuspendLayout();

            foreach (var info in _tableInfos) {
                try {
                    DataTable dt = DataManager.GetTableData(DbName, info.TableName, "", "", "");
                    
                    // 若資料表不存在或全空
                    if (dt == null || dt.Columns.Count == 0 || dt.Rows.Count == 0) {
                        info.GBox.Visible = false;
                        info.Dgv.DataSource = null;
                        continue;
                    }

                    DataView dv = dt.DefaultView;
                    List<string> filters = new List<string>();
                    
                    // 根據各別指定的欄位進行過濾
                    if (!string.IsNullOrEmpty(nameKey) && dt.Columns.Contains(info.NameSearchCol)) 
                        filters.Add($"[{info.NameSearchCol}] LIKE '%{nameKey.Replace("'", "''")}%'");
                    
                    if (!string.IsNullOrEmpty(casKey) && dt.Columns.Contains(info.CasSearchCol)) 
                        filters.Add($"[{info.CasSearchCol}] LIKE '%{casKey.Replace("'", "''")}%'");

                    if (filters.Count > 0) 
                        dv.RowFilter = string.Join(" AND ", filters);
                    else
                        dv.RowFilter = ""; // 無關鍵字時顯示全資料
                    
                    DataTable resultTable = dv.ToTable();

                    // 有篩選出資料時才顯示
                    if (resultTable.Rows.Count > 0) {
                        info.Dgv.DataSource = resultTable;
                        if (info.Dgv.Columns.Contains("Id")) info.Dgv.Columns["Id"].Visible = false;

                        // 動態隱藏內容完全空白的欄位，節省版面
                        foreach (DataGridViewColumn col in info.Dgv.Columns) {
                            if (col.Name == "Id") continue;
                            bool hasData = false;
                            foreach (DataRow row in resultTable.Rows) {
                                if (row[col.Name] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col.Name].ToString())) {
                                    hasData = true;
                                    break;
                                }
                            }
                            col.Visible = hasData; 
                        }
                        info.GBox.Visible = true;
                    } else {
                        info.Dgv.DataSource = null;
                        info.GBox.Visible = false;
                    }
                } catch {
                    info.Dgv.DataSource = null;
                    info.GBox.Visible = false;
                }
            }

            // 🟢 恢復 UI 佈局，一次性繪製完成
            _flpResultsContainer.ResumeLayout(true);
        }

        private void ExportToPdf()
        {
            var visibleTables = _tableInfos.Where(t => t.GBox.Visible && t.Dgv.Rows.Count > 0).ToList();

            if (visibleTables.Count == 0) {
                MessageBox.Show("目前暫無搜尋數據可供導出報表。");
                return;
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

                // 報表總表頭
                g.DrawString("化學品法規符核度分析報表", fTitle, Brushes.DarkSlateGray, x, y);
                y += 40;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   台灣玻璃彰濱廠", fBody, Brushes.Gray, x, y);
                y += 35;

                while (currentTableIndex < visibleTables.Count) {
                    var info = visibleTables[currentTableIndex];
                    var visCols = info.Dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    if (visCols.Count == 0) {
                        currentTableIndex++; continue;
                    }

                    // 畫該表的標題和表頭
                    if (currentRowIndex == 0) {
                        if (y + 80 > e.MarginBounds.Bottom) {
                            e.HasMorePages = true; return;
                        }
                        
                        g.DrawString(info.Title + info.ExtraNotice, fSubTitle, Brushes.Teal, x, y);
                        y += 30;

                        float totalW = visCols.Sum(c => c.Width);
                        float scale = e.MarginBounds.Width / totalW;
                        if (scale > 1.2f) scale = 1.2f;

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
                    }

                    // 畫該表的資料列
                    float tableTotalW = visCols.Sum(c => c.Width);
                    float tableScale = e.MarginBounds.Width / tableTotalW;
                    if (tableScale > 1.2f) tableScale = 1.2f;

                    while (currentRowIndex < info.Dgv.Rows.Count) {
                        float rowH = 30;
                        if (y + rowH > e.MarginBounds.Bottom) {
                            e.HasMorePages = true; return;
                        }

                        float currX = x;
                        foreach (var col in visCols) {
                            RectangleF rect = new RectangleF(currX, y, col.Width * tableScale, rowH);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            string val = info.Dgv[col.Index, currentRowIndex].Value?.ToString() ?? "";
                            g.DrawString(val, fBody, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center });
                            currX += col.Width * tableScale;
                        }
                        y += rowH;
                        currentRowIndex++;
                    }

                    // 該表印完，進入下一個表
                    y += 20; 
                    currentRowIndex = 0;
                    currentTableIndex++;
                }
                
                e.HasMorePages = false;
                currentTableIndex = 0;
                currentRowIndex = 0;
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
