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

        // 輔助資料結構
        private class Node
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public RectangleF Bounds { get; set; }
            public int Level { get; set; } = 0; // 用於排版 (0=來源, 1=聚合中介, 2=最終目標)
        }

        private class SyncEdge
        {
            public Node Source { get; set; }
            public Node Target { get; set; }
            public string SyncType { get; set; }
            public string Detail { get; set; }
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

            // 1. 標題區
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 80, BackColor = Color.White };
            pnlHeader.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlHeader.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            Label lblTitle = new Label { 
                Text = "🗺️ 系統動態流程與架構圖", 
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), 
                ForeColor = Color.DarkSlateBlue, 
                Location = new Point(20, 20), 
                AutoSize = true 
            };
            pnlHeader.Controls.Add(lblTitle);
            mainLayout.Controls.Add(pnlHeader, 0, 0);

            // 2. 內容區 (使用 TabControl 分為「視覺化流程圖」與「樹狀選單清單」)
            TabControl tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(15, 10) };

            // --- Tab 1: 動態資料同步關聯圖 ---
            TabPage tabGraph = new TabPage("📊 資料同步寫入流程圖 (動態繪製)");
            tabGraph.BackColor = Color.White;

            _graphPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(250, 250, 252) };
            _graphPanel.Paint += GraphPanel_Paint;
            
            Button btnRefresh = new Button { Text = "🔄 重新整理流程圖", Size = new Size(180, 40), Location = new Point(20, 15), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (s, e) => LoadGraphData();

            _graphPanel.Controls.Add(btnRefresh);
            tabGraph.Controls.Add(_graphPanel);

            // --- Tab 2: 選單與功能讀取總覽 ---
            TabPage tabTree = new TabPage("📋 系統選單與讀取架構總覽");
            tabTree.BackColor = Color.White;
            TreeView tvMenu = BuildMenuTreeView();
            tabTree.Controls.Add(tvMenu);

            tabControl.TabPages.Add(tabGraph);
            tabControl.TabPages.Add(tabTree);

            mainLayout.Controls.Add(tabControl, 0, 1);

            // 初始載入圖表資料
            LoadGraphData();

            return mainLayout;
        }

        // ==========================================
        // 視覺化流程圖邏輯
        // ==========================================
        private void LoadGraphData()
        {
            _nodes.Clear();
            _syncEdges.Clear();

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
                        string syncType = row.Table.Columns.Contains("SyncType") ? row["SyncType"].ToString() : "單向同步";
                        string detail = $"【{row["SrcMatchCol"]}】➡️【{row["TgtMatchCol"]}】";

                        string sKey = $"{sDb}.{sTb}";
                        string tKey = $"{tDb}.{tTb}";

                        if (!_nodes.ContainsKey(sKey)) _nodes[sKey] = new Node { DbName = sDb, TableName = sTb };
                        if (!_nodes.ContainsKey(tKey)) _nodes[tKey] = new Node { DbName = tDb, TableName = tTb };

                        _syncEdges.Add(new SyncEdge
                        {
                            Source = _nodes[sKey],
                            Target = _nodes[tKey],
                            SyncType = syncType,
                            Detail = detail
                        });
                    }

                    CalculateNodeLayout();
                }
            }
            catch { }

            _graphPanel.Invalidate(); // 觸發重繪
        }

        private void CalculateNodeLayout()
        {
            // 簡易排版演算法：依據來源與目標決定層級 (Level)
            foreach (var edge in _syncEdges)
            {
                edge.Target.Level = Math.Max(edge.Target.Level, edge.Source.Level + 1);
            }

            int nodeWidth = 220;
            int nodeHeight = 80;
            int xSpacing = 350;
            int ySpacing = 150;

            var levelGroups = _nodes.Values.GroupBy(n => n.Level).OrderBy(g => g.Key);
            
            int startX = 50;
            int startY = 80;

            foreach (var group in levelGroups)
            {
                int currentY = startY;
                foreach (var node in group)
                {
                    node.Bounds = new RectangleF(startX + (node.Level * xSpacing), currentY, nodeWidth, nodeHeight);
                    currentY += ySpacing;
                }
            }

            // 確保畫布夠大支援捲動
            int maxX = (int)_nodes.Values.Max(n => n.Bounds.Right) + 100;
            int maxY = (int)_nodes.Values.Max(n => n.Bounds.Bottom) + 100;
            _graphPanel.AutoScrollMinSize = new Size(maxX, maxY);
        }

        private void GraphPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            if (_nodes.Count == 0)
            {
                g.DrawString("目前尚未建立任何資料同步與寫入規則。\n您可以前往「設定 > 資料庫設定」中新增同步規則。", 
                             new Font("Microsoft JhengHei UI", 16F), Brushes.Gray, new PointF(50, 100));
                return;
            }

            // 移動原點以配合捲動軸
            g.TranslateTransform(_graphPanel.AutoScrollPosition.X, _graphPanel.AutoScrollPosition.Y);

            // 1. 繪製連線 (Edges)
            using (Pen penSingle = new Pen(Color.SteelBlue, 3) { CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penDouble = new Pen(Color.DarkOrange, 3) { CustomStartCap = new AdjustableArrowCap(5, 5, true), CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            {
                foreach (var edge in _syncEdges)
                {
                    PointF ptStart = new PointF(edge.Source.Bounds.Right, edge.Source.Bounds.Top + edge.Source.Bounds.Height / 2);
                    PointF ptEnd = new PointF(edge.Target.Bounds.Left, edge.Target.Bounds.Top + edge.Target.Bounds.Height / 2);

                    Pen currentPen = edge.SyncType.Contains("雙向") ? penDouble : penSingle;

                    // 繪製貝茲曲線 (Bezier) 讓線條呈現弧度
                    PointF ctrl1 = new PointF(ptStart.X + 100, ptStart.Y);
                    PointF ctrl2 = new PointF(ptEnd.X - 100, ptEnd.Y);
                    g.DrawBezier(currentPen, ptStart, ctrl1, ctrl2, ptEnd);

                    // 繪製同步資訊文字
                    PointF ptText = new PointF((ptStart.X + ptEnd.X) / 2 - 50, (ptStart.Y + ptEnd.Y) / 2 - 25);
                    g.FillRectangle(new SolidBrush(Color.FromArgb(200, 255, 255, 255)), ptText.X, ptText.Y, 150, 20);
                    g.DrawString(edge.Detail, new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Brushes.DimGray, ptText);
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

                // 標題 (資料庫)
                RectangleF rectTitle = new RectangleF(rect.X, rect.Y + 10, rect.Width, 25);
                g.DrawString($"🗄️ {node.DbName}", new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Brushes.SlateGray, rectTitle, new StringFormat { Alignment = StringAlignment.Center });

                // 內容 (資料表)
                RectangleF rectBody = new RectangleF(rect.X, rect.Y + 40, rect.Width, 30);
                g.DrawString(node.TableName, new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Brushes.Black, rectBody, new StringFormat { Alignment = StringAlignment.Center });
            }
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

            TreeNode rootNode = new TreeNode("📱 Safety System 核心選單架構與資料流向") { NodeFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };

            // 建立標準分類
            Dictionary<string, TreeNode> categoryNodes = new Dictionary<string, TreeNode>();
            string[] standardCategories = { "日常作業", "工安", "化學品", "護理", "空污", "水污", "廢棄物", "消防", "檢測數據", "教育訓練", "法規", "ESG", "ISO14001" };

            foreach (string cat in standardCategories)
            {
                TreeNode catNode = new TreeNode($"📁 {cat}") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                categoryNodes[cat] = catNode;
                rootNode.Nodes.Add(catNode);
            }

            // 加入動態選單資料
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

                        TreeNode menuNode = new TreeNode($"📄 {tableName}") { ToolTipText = $"讀取/寫入資料庫：{dbName} -> {tableName}" };

                        if (categoryNodes.ContainsKey(category))
                        {
                            categoryNodes[category].Nodes.Add(menuNode);
                        }
                        else
                        {
                            // 處理個人隱藏選單
                            if (!categoryNodes.ContainsKey(category))
                            {
                                TreeNode newCatNode = new TreeNode($"🔒 {category} (隱藏選單)") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.Crimson };
                                categoryNodes[category] = newCatNode;
                                rootNode.Nodes.Add(newCatNode);
                            }
                            categoryNodes[category].Nodes.Add(menuNode);
                        }
                    }
                }

                // 加入外部應用程式連結
                TreeNode appLinkNode = new TreeNode("🔗 外部應用連結 (AppLinks)") { NodeFont = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                DataTable dtLinks = DataManager.GetTableData("SystemConfig", "AppLinks", "", "", "");
                if (dtLinks != null)
                {
                    foreach (DataRow row in dtLinks.Rows)
                    {
                        appLinkNode.Nodes.Add(new TreeNode($"🚀 {row["選單名稱"]} -> 執行路徑: {row["執行路徑"]}"));
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
