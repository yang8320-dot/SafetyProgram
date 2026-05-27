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
        private ToolTip _toolTip;
        private List<SyncEdge> _syncEdges = new List<SyncEdge>();
        private Dictionary<string, Node> _nodes = new Dictionary<string, Node>();

        private Dictionary<RectangleF, string> _hoverAreas = new Dictionary<RectangleF, string>();
        private RectangleF _currentHoverRect = RectangleF.Empty;

        private enum EdgeCat { CustomSync, SystemSync }

        private class Node
        {
            public string DbName { get; set; }
            public string TableName { get; set; }
            public RectangleF Bounds { get; set; }
            public int Level { get; set; } = 0; 
            public bool HasBackgroundCalc { get; set; } 
            public string CalcDetail { get; set; }
        }

        private class SyncEdge
        {
            public Node Source { get; set; }
            public Node Target { get; set; }
            public string ShortTitle { get; set; } 
            public string DetailText { get; set; } 
            public EdgeCat Category { get; set; }
        }

        public Control GetView()
        {
            _toolTip = new ToolTip { AutoPopDelay = 10000, InitialDelay = 200, ReshowDelay = 100, UseAnimation = true, UseFading = true };

            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                RowCount = 2,
                ColumnCount = 1
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 1. 標題區 (移除所有 Emoji 避免方塊亂碼)
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 90, BackColor = Color.White };
            pnlHeader.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlHeader.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            Label lblTitle = new Label { 
                Text = "系統動態流程與架構圖", 
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), 
                ForeColor = Color.DarkSlateBlue, 
                Location = new Point(20, 15), 
                AutoSize = true 
            };
            
            Label lblLegend = new Label {
                Text = "提示：滑鼠停留在【連線標籤】或【背景運算】上，可查看詳細轉換規則。    [藍色] 自訂單向    [橘色] 自訂雙向    [紫色] 系統強制聚合",
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                ForeColor = Color.DimGray,
                Location = new Point(25, 55),
                AutoSize = true
            };

            pnlHeader.Controls.Add(lblTitle);
            pnlHeader.Controls.Add(lblLegend);
            mainLayout.Controls.Add(pnlHeader, 0, 0);

            // 2. 內容區
            TabControl tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(15, 8) };

            TabPage tabGraph = new TabPage("資料轉換與寫入流程圖 (動態繪製)");
            tabGraph.BackColor = Color.White;

            _graphPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(250, 250, 252) };
            _graphPanel.Paint += GraphPanel_Paint;
            _graphPanel.MouseMove += GraphPanel_MouseMove;
            
            Button btnRefresh = new Button { Text = "重新整理流程圖", Size = new Size(150, 40), Location = new Point(20, 15), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (s, e) => LoadGraphData();

            _graphPanel.Controls.Add(btnRefresh);
            tabGraph.Controls.Add(_graphPanel);

            TabPage tabTree = new TabPage("系統實體資料庫與資料表總覽");
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
        // 資料讀取與節點建構
        // ==========================================
        private void LoadGraphData()
        {
            _nodes.Clear();
            _syncEdges.Clear();

            // 1. 讀取使用者自訂同步規則 (SyncRules)
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
                        if (sMatch != tMatch && sMatch.Contains("日") && tMatch.Contains("月")) {
                            transformInfo = "\n※轉換: 截取字串前7碼(YYYY-MM)並SUM加總";
                        }

                        string shortTitle = syncType.Contains("雙向") ? "自訂雙向" : "自訂單向";
                        string detail = $"來源: [{sMatch}] ➡️ 目標: [{tMatch}]\n寫入: [{sSync}] ➡️ [{tSync}]{transformInfo}";

                        AddEdge(sDb, sTb, tDb, tTb, shortTitle, detail, EdgeCat.CustomSync);
                    }
                }
            }
            catch { }

            // 2. 系統強制聚合流程
            string[] lawTables = { "環保法規", "職安衛法規", "消防法規", "其它法規" };
            foreach(var lawTb in lawTables)
            {
                AddEdge("法規", lawTb, "法規", "法規目錄一覽", "系統聚合", 
                        $"掃描同法規名稱進行歸類聚合\n提取最新日期與最高適用性回填", EdgeCat.SystemSync);
            }

            // 3. 系統內部運算
            string[] waterTables = { "WaterMeterReadings", "WaterUsageDaily" };
            foreach(var wTb in waterTables)
            {
                EnsureNode("Water", wTb).HasBackgroundCalc = true;
                EnsureNode("Water", wTb).CalcDetail = "存檔時自動尋找下一筆紀錄\n執行 (Next - Target) 差值運算\n產生(日/月統計)並回填入原資料列";
            }

            CalculateNodeLayout();
            _graphPanel.Invalidate(); 
        }

        private Node EnsureNode(string db, string tb)
        {
            string key = $"{db}.{tb}";
            if (!_nodes.ContainsKey(key)) _nodes[key] = new Node { DbName = db, TableName = tb };
            return _nodes[key];
        }

        private void AddEdge(string sDb, string sTb, string tDb, string tTb, string shortTitle, string detail, EdgeCat cat)
        {
            Node sourceNode = EnsureNode(sDb, sTb);
            Node targetNode = EnsureNode(tDb, tTb);

            _syncEdges.Add(new SyncEdge
            {
                Source = sourceNode,
                Target = targetNode,
                ShortTitle = shortTitle,
                DetailText = detail,
                Category = cat
            });
        }

        // ==========================================
        // 緊湊排版引擎 (解決文字重疊)
        // ==========================================
        private void CalculateNodeLayout()
        {
            foreach (var n in _nodes.Values) n.Level = 0;

            bool changed;
            int safeguard = 0;
            do {
                changed = false;
                foreach (var edge in _syncEdges) {
                    if (edge.Target.Level <= edge.Source.Level) {
                        edge.Target.Level = edge.Source.Level + 1;
                        changed = true;
                    }
                }
                safeguard++;
            } while (changed && safeguard < 100);

            // 🌟 極致優化：增加 Y 軸間距，讓多行文字絕對不會互相重疊
            int nodeWidth = 200;
            int nodeHeight = 70;
            int xSpacing = 380; 
            int ySpacing = 160; // 將高度間距拉開，文字框高度約 60px，160 能完美錯開

            var levelGroups = _nodes.Values.GroupBy(n => n.Level).OrderBy(g => g.Key);
            
            int startX = 40;
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

            int maxX = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Right) + 200 : 800;
            int maxY = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Bottom) + 200 : 600;
            _graphPanel.AutoScrollMinSize = new Size(maxX, maxY);
        }

        // ==========================================
        // 繪圖引擎
        // ==========================================
        private void GraphPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            _hoverAreas.Clear(); 

            if (_nodes.Count == 0) return;

            g.TranslateTransform(_graphPanel.AutoScrollPosition.X, _graphPanel.AutoScrollPosition.Y);

            Font textFont = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

            // 1. 繪製連線
            using (Pen penCustomSingle = new Pen(Color.SteelBlue, 2) { CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penCustomDouble = new Pen(Color.DarkOrange, 2) { CustomStartCap = new AdjustableArrowCap(5, 5, true), CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            using (Pen penSysSync = new Pen(Color.MediumPurple, 2) { CustomEndCap = new AdjustableArrowCap(5, 5, true) })
            {
                foreach (var edge in _syncEdges)
                {
                    Pen currentPen = edge.Category == EdgeCat.SystemSync ? penSysSync : 
                                     (edge.ShortTitle.Contains("雙向") ? penCustomDouble : penCustomSingle);

                    PointF ptStart = new PointF(edge.Source.Bounds.Right, edge.Source.Bounds.Top + edge.Source.Bounds.Height / 2);
                    PointF ptEnd = new PointF(edge.Target.Bounds.Left, edge.Target.Bounds.Top + edge.Target.Bounds.Height / 2);

                    float offset = 100;
                    PointF ctrl1 = new PointF(ptStart.X + offset, ptStart.Y);
                    PointF ctrl2 = new PointF(ptEnd.X - offset, ptEnd.Y);

                    g.DrawBezier(currentPen, ptStart, ctrl1, ctrl2, ptEnd);

                    // 繪製短標籤 (不再繪製長篇大論)
                    SizeF tSize = g.MeasureString(edge.ShortTitle, textFont);
                    PointF ptText = new PointF(ptStart.X + 20, ptStart.Y - 12);
                    RectangleF bgRect = new RectangleF(ptText.X, ptText.Y, tSize.Width + 6, tSize.Height + 4);
                    
                    g.FillRectangle(new SolidBrush(Color.FromArgb(245, 255, 255, 255)), bgRect);
                    g.DrawRectangle(Pens.DarkGray, bgRect.X, bgRect.Y, bgRect.Width, bgRect.Height);
                    
                    Brush fBrush = edge.Category == EdgeCat.SystemSync ? Brushes.Purple : Brushes.DarkSlateGray;
                    g.DrawString(edge.ShortTitle, textFont, fBrush, new PointF(ptText.X + 3, ptText.Y + 2));

                    // 將長篇文字塞入隱形熱區，等待滑鼠懸停
                    RectangleF hitRect = new RectangleF(bgRect.X + _graphPanel.AutoScrollPosition.X, bgRect.Y + _graphPanel.AutoScrollPosition.Y, bgRect.Width, bgRect.Height);
                    _hoverAreas[hitRect] = edge.DetailText;
                }
            }

            // 2. 繪製節點
            foreach (var node in _nodes.Values)
            {
                Rectangle rect = Rectangle.Round(node.Bounds);
                
                using (GraphicsPath path = GetRoundedRectPath(rect, 10))
                using (LinearGradientBrush brush = new LinearGradientBrush(rect, Color.White, Color.FromArgb(240, 245, 250), LinearGradientMode.Vertical))
                using (Pen borderPen = new Pen(Color.FromArgb(60, 80, 100), 1))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(borderPen, path);
                }

                RectangleF rectTitle = new RectangleF(rect.X, rect.Y + 8, rect.Width, 20);
                g.DrawString($"[{node.DbName}]", new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Brushes.SlateGray, rectTitle, new StringFormat { Alignment = StringAlignment.Center });

                RectangleF rectBody = new RectangleF(rect.X, rect.Y + 30, rect.Width, 25);
                g.DrawString(node.TableName, new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Brushes.Black, rectBody, new StringFormat { Alignment = StringAlignment.Center });

                if (node.HasBackgroundCalc)
                {
                    string badgeText = "[背景運算]";
                    SizeF bSize = g.MeasureString(badgeText, textFont);
                    RectangleF badgeRect = new RectangleF(rect.Right - bSize.Width - 10, rect.Bottom - bSize.Height - 5, bSize.Width + 6, bSize.Height + 2);
                    
                    using (GraphicsPath bPath = GetRoundedRectPath(Rectangle.Round(badgeRect), 5))
                    {
                        g.FillPath(Brushes.SeaGreen, bPath);
                        g.DrawString(badgeText, textFont, Brushes.White, new PointF(badgeRect.X + 3, badgeRect.Y + 1));
                    }

                    RectangleF hitRect = new RectangleF(badgeRect.X + _graphPanel.AutoScrollPosition.X, badgeRect.Y + _graphPanel.AutoScrollPosition.Y, badgeRect.Width, badgeRect.Height);
                    _hoverAreas[hitRect] = node.CalcDetail;
                }
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

        private void GraphPanel_MouseMove(object sender, MouseEventArgs e)
        {
            Point mousePt = e.Location;
            bool isHit = false;

            foreach (var kvp in _hoverAreas)
            {
                if (kvp.Key.Contains(mousePt))
                {
                    isHit = true;
                    _graphPanel.Cursor = Cursors.Help;

                    if (_currentHoverRect != kvp.Key)
                    {
                        _currentHoverRect = kvp.Key;
                        _toolTip.Show(kvp.Value, _graphPanel, mousePt.X + 15, mousePt.Y + 15);
                    }
                    break;
                }
            }

            if (!isHit)
            {
                _graphPanel.Cursor = Cursors.Default;
                if (_currentHoverRect != RectangleF.Empty)
                {
                    _toolTip.Hide(_graphPanel);
                    _currentHoverRect = RectangleF.Empty;
                }
            }
        }

        // ==========================================
        // 樹狀選單 (完整對應實體庫與資料表)
        // ==========================================
        private TreeView BuildMenuTreeView()
        {
            // 🌟 修正：取消 NodeFont 強制覆寫，改為在 TreeView 全域設定 Font，解決標題被切斷的 Bug
            TreeView tv = new TreeView 
            { 
                Dock = DockStyle.Fill, 
                BorderStyle = BorderStyle.None, 
                ItemHeight = 28, 
                Margin = new Padding(10),
                Font = new Font("Microsoft JhengHei UI", 12F)
            };
            
            TreeNode rootNode = new TreeNode("Safety System 實體資料庫與對應資料表總覽");

            try
            {
                // 讀取 App_DbConfig 中的實體資料庫結構
                var dbMap = App_DbConfig.GetDbMapCache();

                foreach (var kvp in dbMap)
                {
                    string dbEnName = kvp.Key;
                    string dbChName = kvp.Value.ChDbName;

                    TreeNode dbNode = new TreeNode($"[資料庫] {dbChName} ({dbEnName}.sqlite)") { ForeColor = Color.DarkSlateBlue };

                    foreach (var tb in kvp.Value.Tables)
                    {
                        string tbEnName = tb.Key;
                        string tbChName = tb.Value;
                        TreeNode tbNode = new TreeNode($"- {tbChName} ({tbEnName})");
                        dbNode.Nodes.Add(tbNode);
                    }
                    rootNode.Nodes.Add(dbNode);
                }

                // 加入外部應用程式連結
                TreeNode appLinkNode = new TreeNode("[外掛] 外部應用連結 (AppLinks)") { ForeColor = Color.DarkOliveGreen };
                DataTable dtLinks = DataManager.GetTableData("SystemConfig", "AppLinks", "", "", "");
                if (dtLinks != null)
                {
                    foreach (DataRow row in dtLinks.Rows)
                    {
                        appLinkNode.Nodes.Add(new TreeNode($"> {row["選單名稱"]} -> 路徑: {row["執行路徑"]}"));
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
