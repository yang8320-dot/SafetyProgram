/// FILE: Safety_System/settings/App_SystemFlowchart.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SystemFlowchart
    {
        private Panel _graphPanel;
        private List<SyncEdge> _syncEdges = new List<SyncEdge>();
        private Dictionary<string, Node> _nodes = new Dictionary<string, Node>();

        private enum EdgeCat { CustomSync, SystemSync, SystemCalc }

        // 輔助資料結構
        private class Node
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public RectangleF Bounds { get; set; }
            public int Level { get; set; } = 0; 
        }

        private class SyncEdge
        {
            public Node Source { get; set; }
            public Node Target { get; set; }
            public string SyncType { get; set; }
            public string Detail { get; set; }
            public EdgeCat Category { get; set; }
        }

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                RowCount = 2,
                ColumnCount = 1
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 1. 標題區與圖例說明
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 100, BackColor = Color.White };
            pnlHeader.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlHeader.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            Label lblTitle = new Label { 
                Text = "系統動態流程與架構圖", 
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), 
                ForeColor = Color.DarkSlateBlue, 
                Location = new Point(20, 15), 
                AutoSize = true 
            };
            
            // 修正：移除作業系統可能不支援的 Emoji
            Label lblLegend = new Label {
                Text = "圖例說明： [藍色實線] 自訂單向同步    [橘色實線] 自訂雙向同步    [紫色實線] 系統強制聚合寫入    [綠色虛線] 系統內部運算",
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                ForeColor = Color.DimGray,
                Location = new Point(25, 65),
                AutoSize = true
            };

            pnlHeader.Controls.Add(lblTitle);
            pnlHeader.Controls.Add(lblLegend);
            mainLayout.Controls.Add(pnlHeader, 0, 0);

            // 2. 內容區
            TabControl tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(15, 10) };

            TabPage tabGraph = new TabPage("資料轉換與寫入流程圖 (動態繪製)");
            tabGraph.BackColor = Color.White;

            _graphPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(250, 250, 252) };
            _graphPanel.Paint += GraphPanel_Paint;
            
            Button btnRefresh = new Button { Text = "重新整理流程圖", Size = new Size(180, 40), Location = new Point(20, 15), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (s, e) => LoadGraphData();

            _graphPanel.Controls.Add(btnRefresh);
            tabGraph.Controls.Add(_graphPanel);

            TabPage tabTree = new TabPage("系統選單與實體資料庫總覽");
            tabTree.BackColor = Color.White;
            TreeView tvMenu = BuildMenuTreeView();
            tabTree.Controls.Add(tvMenu);

            tabControl.TabPages.Add(tabGraph);
            tabControl.TabPages.Add(tabTree);

            mainLayout.Controls.Add(tabControl, 0, 1);

            LoadGraphData();

            return mainLayout;
        }

        // ==========================================
        // 視覺化流程圖與關聯建立邏輯
        // ==========================================
        private void LoadGraphData()
        {
            _nodes.Clear();
            _syncEdges.Clear();

            // ----------------------------------------------------
            // 1. 讀取使用者自定義的動態同步規則 (SyncRules)
            // ----------------------------------------------------
            try
            {
                DataTable dtSync = DataManager.GetTableData("SystemConfig", "SyncRules", "", "", "");
                if (dtSync != null && dtSync.Rows.Count > 0)
                {
                    foreach (DataRow row in dtSync.Rows)
                    {
                        string sDb = row["SrcDb"].ToString();
                        string sTb = row["SrcTable"].ToString();
                        string tDb = row["TgtDb"].ToString();
                        string tTb = row["TgtTable"].ToString();
                        string sMatch = row["SrcMatchCol"].ToString();
                        string tMatch = row["TgtMatchCol"].ToString();
                        string sSync = row["SrcSyncCol"].ToString();
                        string tSync = row["TgtSyncCol"].ToString();
                        string syncType = row.Table.Columns.Contains("SyncType") ? row["SyncType"].ToString() : "單向同步";

                        string transformInfo = "";
                        if (sMatch != tMatch && sMatch.Contains("日") && tMatch.Contains("月"))
                        {
                            transformInfo = "\n※資料轉換: 截取字串前7碼(YYYY-MM)\n並將來源數值執行 SUM 聚合加總";
                        }

                        string detail = $"對應條件: [{sMatch}] ➡️ [{tMatch}]\n寫入目標: [{sSync}] ➡️ [{tSync}]{transformInfo}";

                        AddEdge(sDb, sTb, tDb, tTb, syncType, detail, EdgeCat.CustomSync);
                    }
                }
            }
            catch { }

            // ----------------------------------------------------
            // 2. 加入系統內建強制流程
            // ----------------------------------------------------
            string[] lawTables = { "環保法規", "職安衛法規", "消防法規", "其它法規" };
            foreach(var lawTb in lawTables)
            {
                AddEdge("法規", lawTb, "法規", "法規目錄一覽", "系統單向聚合", 
                        $"※資料轉換: 掃描 [{lawTb}]\n將相同法規名稱歸類聚合\n取最新日期與最高適用性寫入目錄", EdgeCat.SystemSync);
            }

            // ----------------------------------------------------
            // 3. 加入系統內部運算攔截
            // ----------------------------------------------------
            string[] waterTables = { "WaterMeterReadings", "WaterUsageDaily" };
            foreach(var wTb in waterTables)
            {
                AddEdge("Water", wTb, "Water", wTb, "系統背景運算", 
                        $"※資料計算: 存檔時尋找下一筆紀錄\n執行 (Next - Target) 差值運算\n產生(日統計)回填入原資料列", EdgeCat.SystemCalc);
            }

            CalculateNodeLayout();
            _graphPanel.Invalidate(); 
        }

        private void AddEdge(string sDb, string sTb, string tDb, string tTb, string type, string detail, EdgeCat cat)
        {
            string sKey = $"{sDb}.{sTb}";
            string tKey = $"{tDb}.{tTb}";

            if (!_nodes.ContainsKey(sKey)) _nodes[sKey] = new Node { DbName = sDb, TableName = sTb };
            if (!_nodes.ContainsKey(tKey)) _nodes[tKey] = new Node { DbName = tDb, TableName = tTb };

            _syncEdges.Add(new SyncEdge
            {
                Source = _nodes[sKey],
                Target = _nodes[tKey],
                SyncType = type,
                Detail = detail,
                Category = cat
            });
        }

        private void CalculateNodeLayout()
        {
            foreach (var n in _nodes.Values) n.Level = 0;

            bool changed;
            int safeguard = 0;
            do {
                changed = false;
                foreach (var edge in _syncEdges) {
                    if (edge.Source != edge.Target && edge.Target.Level <= edge.Source.Level) {
                        edge.Target.Level = edge.Source.Level + 1;
                        changed = true;
                    }
                }
                safeguard++;
            } while (changed && safeguard < 100);

            // 🌟 修正點 1：大幅增加 X 與 Y 的節點間距，給予文字更多擺放空間
            int nodeWidth = 240;
            int nodeHeight = 85;
            int xSpacing = 550; // 原 450 -> 改 550
            int ySpacing = 220; // 原 160 -> 改 220，避免上下行文字打架

            var levelGroups = _nodes.Values.GroupBy(n => n.Level).OrderBy(g => g.Key);
            
            int startX = 50;
            int startY = 100; 

            foreach (var group in levelGroups)
            {
                int currentY = startY;
                foreach (var node in group)
                {
                    node.Bounds = new RectangleF(startX + (node.Level * xSpacing), currentY, nodeWidth, nodeHeight);
                    currentY += ySpacing;
                }
            }

            int maxX = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Right) + 300 : 1000;
            int maxY = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Bottom) + 300 : 800;
            _graphPanel.AutoScrollMinSize = new Size(maxX, maxY);
        }

        private void GraphPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            if (_nodes.Count == 0)
            {
                g.DrawString("目前沒有任何資料流動與同步規則。", new Font("Microsoft JhengHei UI", 16F), Brushes.Gray, new PointF(50, 100));
                return;
            }

            g.TranslateTransform(_graphPanel.AutoScrollPosition.X, _graphPanel.AutoScrollPosition.Y);

            Font textFont = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
            Brush textBrush = Brushes.DarkSlateGray;

            // 1. 繪製連線 (Edges)
            using (Pen penCustomSingle = new Pen(Color.SteelBlue, 3) { CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penCustomDouble = new Pen(Color.DarkOrange, 3) { CustomStartCap = new AdjustableArrowCap(5, 5, true), CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penSysSync = new Pen(Color.MediumPurple, 3) { CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penSysCalc = new Pen(Color.SeaGreen, 3) { DashStyle = DashStyle.Dash, CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            {
                foreach (var edge in _syncEdges)
                {
                    Pen currentPen;
                    if (edge.Category == EdgeCat.SystemCalc) currentPen = penSysCalc;
                    else if (edge.Category == EdgeCat.SystemSync) currentPen = penSysSync;
                    else currentPen = edge.SyncType.Contains("雙向") ? penCustomDouble : penCustomSingle;

                    if (edge.Source == edge.Target)
                    {
                        // 自我迴圈 (表示內部運算)
                        PointF ptTop = new PointF(edge.Source.Bounds.Left + edge.Source.Bounds.Width / 2, edge.Source.Bounds.Top);
                        PointF ctrl1 = new PointF(ptTop.X - 100, ptTop.Y - 150);
                        PointF ctrl2 = new PointF(ptTop.X + 100, ptTop.Y - 150);

                        g.DrawBezier(currentPen, ptTop, ctrl1, ctrl2, ptTop);

                        PointF ptText = new PointF(ptTop.X - 140, ptTop.Y - 120);
                        DrawEdgeText(g, $"【{edge.SyncType}】\n{edge.Detail}", textFont, Brushes.DarkGreen, ptText);
                    }
                    else
                    {
                        // 跨表連線 (左到右)
                        PointF ptStart = new PointF(edge.Source.Bounds.Right, edge.Source.Bounds.Top + edge.Source.Bounds.Height / 2);
                        PointF ptEnd = new PointF(edge.Target.Bounds.Left, edge.Target.Bounds.Top + edge.Target.Bounds.Height / 2);

                        float offset = 200; // 固定控制點，讓線條柔和
                        PointF ctrl1 = new PointF(ptStart.X + offset, ptStart.Y);
                        PointF ctrl2 = new PointF(ptEnd.X - offset, ptEnd.Y);

                        g.DrawBezier(currentPen, ptStart, ctrl1, ctrl2, ptEnd);

                        // 🌟 修正點 2：解決文字重疊的終極解法
                        // 將文字錨點強制綁定在「來源節點 (Source) 右側稍微出來一點」的地方
                        // 因為每個來源節點的高度 (Y軸) 都不同，這樣就算所有線終點都一樣，文字也絕對不會上下撞在一起。
                        PointF ptText = new PointF(ptStart.X + 30, ptStart.Y - 20);
                        
                        Brush fBrush = edge.Category == EdgeCat.SystemSync ? Brushes.Purple : textBrush;
                        DrawEdgeText(g, $"【{edge.SyncType}】\n{edge.Detail}", textFont, fBrush, ptText);
                    }
                }
            }

            // 2. 繪製節點 (Nodes)
            foreach (var node in _nodes.Values)
            {
                Rectangle rect = Rectangle.Round(node.Bounds);
                int cornerRadius = 15;

                using (GraphicsPath path = GetRoundedRectPath(rect, cornerRadius))
                using (LinearGradientBrush brush = new LinearGradientBrush(rect, Color.White, Color.FromArgb(235, 240, 245), LinearGradientMode.Vertical))
                using (Pen borderPen = new Pen(Color.FromArgb(45, 62, 80), 2))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(borderPen, path);
                }

                // 修正：移除 Emoji，改用純文字標籤
                RectangleF rectTitle = new RectangleF(rect.X, rect.Y + 12, rect.Width, 25);
                g.DrawString($"[庫] {node.DbName}", new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Brushes.SlateGray, rectTitle, new StringFormat { Alignment = StringAlignment.Center });

                RectangleF rectBody = new RectangleF(rect.X, rect.Y + 42, rect.Width, 30);
                g.DrawString(node.TableName, new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Brushes.Black, rectBody, new StringFormat { Alignment = StringAlignment.Center });
            }
        }

        private void DrawEdgeText(Graphics g, string text, Font font, Brush textBrush, PointF location)
        {
            SizeF textSize = g.MeasureString(text, font);
            RectangleF bgRect = new RectangleF(location.X, location.Y, textSize.Width + 10, textSize.Height + 6);
            
            // 繪製半透明白底，讓文字壓在線上時能看清楚
            g.FillRectangle(new SolidBrush(Color.FromArgb(230, 255, 255, 255)), bgRect);
            g.DrawRectangle(Pens.LightGray, bgRect.X, bgRect.Y, bgRect.Width, bgRect.Height);
            g.DrawString(text, font, textBrush, new PointF(location.X + 5, location.Y + 3));
        }

        private GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        // ==========================================
        // 樹狀選單與架構清單邏輯
        // ==========================================
        private TreeView BuildMenuTreeView()
        {
            TreeView tv = new TreeView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                ShowNodeToolTips = true,
                ItemHeight = 35,
                Margin = new Padding(10)
            };

            // 修正：移除 Emoji
            TreeNode rootNode = new TreeNode("Safety System 核心選單架構與實體資料庫對應關係") { NodeFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };

            Dictionary<string, TreeNode> categoryNodes = new Dictionary<string, TreeNode>();
            string[] standardCategories = { "日常作業", "工安", "化學品", "護理", "空污", "水污", "廢棄物", "消防", "檢測數據", "教育訓練", "法規", "ESG", "ISO14001" };

            foreach (string cat in standardCategories)
            {
                TreeNode catNode = new TreeNode($"[分類] {cat}") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                categoryNodes[cat] = catNode;
                rootNode.Nodes.Add(catNode);
            }

            try
            {
                DataTable dtMenus = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
                if (dtMenus != null)
                {
                    foreach (DataRow row in dtMenus.Rows)
                    {
                        string category = row["分類"].ToString();
                        string dbName = row["資料庫名"].ToString();
                        string tableName = row["資料表名"].ToString();

                        TreeNode menuNode = new TreeNode($"- {tableName}") { ToolTipText = $"點擊選單將操作實體資料庫：{dbName}.sqlite -> 資料表 [{tableName}]" };

                        if (categoryNodes.ContainsKey(category))
                        {
                            categoryNodes[category].Nodes.Add(menuNode);
                        }
                        else
                        {
                            if (!categoryNodes.ContainsKey(category))
                            {
                                TreeNode newCatNode = new TreeNode($"[隱藏分類] {category}") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.Crimson };
                                categoryNodes[category] = newCatNode;
                                rootNode.Nodes.Add(newCatNode);
                            }
                            categoryNodes[category].Nodes.Add(menuNode);
                        }
                    }
                }

                TreeNode appLinkNode = new TreeNode("[外掛] 外部應用連結 (AppLinks)") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                DataTable dtLinks = DataManager.GetTableData("SystemConfig", "AppLinks", "", "", "");
                if (dtLinks != null)
                {
                    foreach (DataRow row in dtLinks.Rows)
                    {
                        appLinkNode.Nodes.Add(new TreeNode($"> {row["選單名稱"]} -> 實體路徑: {row["執行路徑"]}"));
                    }
                }
                rootNode.Nodes.Add(appLinkNode);
            }
            catch { }

            tv.Nodes.Add(rootNode);
            rootNode.ExpandAll();

            return tv;
        }
    }
}
