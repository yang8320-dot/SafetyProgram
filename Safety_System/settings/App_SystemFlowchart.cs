/// FILE: Safety_System/settings/App_SystemFlowchart.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
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
        
        // 🟢 紀錄已解鎖的隱藏選單
        private HashSet<string> _unlockedMenus = new HashSet<string>();

        private enum EdgeCat { CustomSync, SystemSync }

        private class Node
        {
            public string SystemDbName { get; set; }
            public string SystemTableName { get; set; }
            public string ChDbName { get; set; }   
            public string ChTableName { get; set; } 
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

            TabControl tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(15, 8) };

            TabPage tabGraph = new TabPage("資料轉換與寫入流程圖 (動態繪製)");
            tabGraph.BackColor = Color.White;

            _graphPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(250, 250, 252) };
            _graphPanel.Paint += GraphPanel_Paint;
            _graphPanel.MouseMove += GraphPanel_MouseMove;
            
            Button btnRefresh = new Button { Text = "重新整理流程圖", Size = new Size(180, 40), Location = new Point(20, 15), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
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

        private void LoadGraphData()
        {
            _nodes.Clear();
            _syncEdges.Clear();

            try
            {
                DataTable dtSync = new DataTable();
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM SyncRules", conn))
                    using (var da = new SQLiteDataAdapter(cmd))
                    {
                        da.Fill(dtSync);
                    }
                }

                if (dtSync.Rows.Count > 0)
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

            string[] lawTables = { "環保法規", "職安衛法規", "消防法規", "其它法規" };
            foreach(var lawTb in lawTables)
            {
                AddEdge("法規", lawTb, "法規", "法規目錄一覽", "系統聚合", 
                        $"掃描同法規名稱進行歸類聚合\n提取最新日期與最高適用性回填", EdgeCat.SystemSync);
            }

            string[] waterTables = { "WaterMeterReadings", "WaterUsageDaily" };
            foreach(var wTb in waterTables)
            {
                Node wNode = EnsureNode("Water", wTb);
                wNode.HasBackgroundCalc = true;
                wNode.CalcDetail = "存檔時自動尋找下一筆紀錄\n執行 (Next - Target) 差值運算\n產生(日/月統計)並回填入原資料列";
            }

            CalculateNodeLayout();
            _graphPanel.Invalidate(); 
        }

        private Node EnsureNode(string db, string tb)
        {
            string key = $"{db}.{tb}";
            if (!_nodes.ContainsKey(key)) 
            {
                string chDb = db;
                string chTb = tb;

                var dbMap = App_DbConfig.GetDbMapCache();
                if (dbMap.ContainsKey(db)) {
                    chDb = dbMap[db].ChDbName;
                    if (dbMap[db].Tables.ContainsKey(tb)) {
                        chTb = dbMap[db].Tables[tb];
                    }
                }

                _nodes[key] = new Node { 
                    SystemDbName = db, 
                    SystemTableName = tb,
                    ChDbName = chDb,
                    ChTableName = chTb
                };
            }
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

            int nodeWidth = 260; 
            int nodeHeight = 75; 
            int xSpacing = 400;  
            int ySpacing = 110;  

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

            int maxX = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Right) + 150 : 800;
            int maxY = _nodes.Count > 0 ? (int)_nodes.Values.Max(n => n.Bounds.Bottom) + 150 : 600;
            _graphPanel.AutoScrollMinSize = new Size(maxX, maxY);
        }

        private void GraphPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            _hoverAreas.Clear(); 

            if (_nodes.Count == 0) return;

            g.TranslateTransform(_graphPanel.AutoScrollPosition.X, _graphPanel.AutoScrollPosition.Y);

            Font textFont = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

            // 🟢 圖層 1 (底層)：只負責畫線 (確保線條永遠在最底下)
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

                    float midX = ptStart.X + (ptEnd.X - ptStart.X) / 2;
                    PointF[] points = {
                        ptStart,
                        new PointF(midX, ptStart.Y),
                        new PointF(midX, ptEnd.Y),
                        ptEnd
                    };

                    g.DrawLines(currentPen, points);
                }
            }

            // 🟢 圖層 2 (中層)：畫出所有的說明圖塊 (蓋在線條上方，並帶有部分透明度)
            foreach (var edge in _syncEdges)
            {
                PointF ptStart = new PointF(edge.Source.Bounds.Right, edge.Source.Bounds.Top + edge.Source.Bounds.Height / 2);
                SizeF tSize = g.MeasureString(edge.ShortTitle, textFont);
                PointF ptText = new PointF(ptStart.X + 10, ptStart.Y - 18);
                RectangleF bgRect = new RectangleF(ptText.X, ptText.Y, tSize.Width + 6, tSize.Height + 4);
                
                // 🟢 使用 Alpha 220 呈現微透效果，讓底下的線隱約可見，但不會干擾閱讀
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                {
                    g.FillRectangle(bgBrush, bgRect);
                }
                
                g.DrawRectangle(Pens.DarkGray, bgRect.X, bgRect.Y, bgRect.Width, bgRect.Height);
                
                Brush fBrush = edge.Category == EdgeCat.SystemSync ? Brushes.Purple : Brushes.DarkSlateGray;
                g.DrawString(edge.ShortTitle, textFont, fBrush, new PointF(ptText.X + 3, ptText.Y + 2));

                // 註冊 Hover 區域
                RectangleF hitRect = new RectangleF(bgRect.X + _graphPanel.AutoScrollPosition.X, bgRect.Y + _graphPanel.AutoScrollPosition.Y, bgRect.Width, bgRect.Height);
                _hoverAreas[hitRect] = edge.DetailText;
            }

            // 🟢 圖層 3 (頂層)：畫所有的節點框塊
            StringFormat sfTitle = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
            StringFormat sfBody = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap };

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

                RectangleF rectTitle = new RectangleF(rect.X + 5, rect.Y + 8, rect.Width - 10, 20);
                g.DrawString($"[{node.ChDbName}]", new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Brushes.SlateGray, rectTitle, sfTitle);

                RectangleF rectBody = new RectangleF(rect.X + 5, rect.Y + 30, rect.Width - 10, 30);
                g.DrawString(node.ChTableName, new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Brushes.Black, rectBody, sfBody);

                if (node.HasBackgroundCalc)
                {
                    string badgeText = "[背景運算]";
                    SizeF bSize = g.MeasureString(badgeText, textFont);
                    
                    RectangleF badgeRect = new RectangleF(rect.Right - bSize.Width - 8, rect.Bottom - bSize.Height - 4, bSize.Width + 4, bSize.Height + 2);
                    
                    using (GraphicsPath bPath = GetRoundedRectPath(Rectangle.Round(badgeRect), 4))
                    {
                        g.FillPath(Brushes.SeaGreen, bPath);
                        g.DrawString(badgeText, textFont, Brushes.White, new PointF(badgeRect.X + 2, badgeRect.Y + 1));
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

        private TreeView BuildMenuTreeView()
        {
            TreeView tv = new TreeView 
            { 
                Dock = DockStyle.Fill, 
                BorderStyle = BorderStyle.None, 
                ItemHeight = 28, 
                Margin = new Padding(10),
                Font = new Font("Microsoft JhengHei UI", 12F)
            };
            
            // 🟢 綁定展開前事件，處理隱藏選單密碼驗證
            tv.BeforeExpand += TvMenu_BeforeExpand;

            TreeNode rootNode = new TreeNode("Safety System 實體資料庫與對應資料表總覽");

            try
            {
                var dbMap = App_DbConfig.GetDbMapCache();

                foreach (var kvp in dbMap)
                {
                    string dbEnName = kvp.Key;
                    string dbChName = kvp.Value.ChDbName;

                    TreeNode dbNode = new TreeNode($"[資料庫] {dbChName} ({dbEnName}.sqlite)") { ForeColor = Color.DarkSlateBlue };

                    // 🟢 判斷並設定隱藏選單的標籤 (Tag)
                    if (dbEnName.StartsWith("Menu") && dbEnName.EndsWith("DB"))
                    {
                        string menuName = "";
                        if (dbEnName == "Menu1DB") menuName = "選單1";
                        if (dbEnName == "Menu2DB") menuName = "選單2";
                        if (dbEnName == "Menu3DB") menuName = "選單3";
                        if (dbEnName == "Menu4DB") menuName = "選單4";
                        
                        dbNode.Tag = "HiddenMenu|" + menuName;
                    }

                    foreach (var tb in kvp.Value.Tables)
                    {
                        string tbEnName = tb.Key;
                        string tbChName = tb.Value;
                        TreeNode tbNode = new TreeNode($"- {tbChName} ({tbEnName})");
                        dbNode.Nodes.Add(tbNode);
                    }
                    rootNode.Nodes.Add(dbNode);
                }

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
            
            // 展開所有節點
            rootNode.ExpandAll();

            // 🟢 將標記為隱藏的選單收合起來
            foreach (TreeNode dbNode in rootNode.Nodes)
            {
                if (dbNode.Tag != null && dbNode.Tag.ToString().StartsWith("HiddenMenu|"))
                {
                    dbNode.Collapse();
                }
            }

            return tv;
        }

        // 🟢 攔截展開事件，進行權限驗證
        private void TvMenu_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.Tag != null && e.Node.Tag.ToString().StartsWith("HiddenMenu|"))
            {
                string menuName = e.Node.Tag.ToString().Split('|')[1];
                
                // 檢查是否已解鎖
                if (_unlockedMenus.Contains(menuName)) return;

                // 檢查是否符合預設擁有者帳號 (免密碼放行)
                string currentUser = Environment.UserName.Trim();
                bool isAuthorized = false;

                if (menuName == "選單1" && (string.Equals(currentUser, "黃忠揚", StringComparison.OrdinalIgnoreCase) || string.Equals(currentUser, "TJ700657", StringComparison.OrdinalIgnoreCase))) isAuthorized = true;
                else if (menuName == "選單2" && string.Equals(currentUser, "TJ700228", StringComparison.OrdinalIgnoreCase)) isAuthorized = true;
                else if (menuName == "選單3" && string.Equals(currentUser, "TJ700533", StringComparison.OrdinalIgnoreCase)) isAuthorized = true;
                else if (menuName == "選單4" && string.Equals(currentUser, "TJ204159", StringComparison.OrdinalIgnoreCase)) isAuthorized = true;

                if (isAuthorized)
                {
                    _unlockedMenus.Add(menuName);
                    return;
                }

                // 需要密碼驗證
                if (!VerifyHiddenMenuPassword(menuName))
                {
                    e.Cancel = true; // 阻止展開
                }
                else
                {
                    _unlockedMenus.Add(menuName);
                }
            }
        }

        // 🟢 密碼驗證方法
        private bool VerifyHiddenMenuPassword(string menuName)
        {
            using (Form p = new Form())
            {
                p.Width = 460; 
                p.Height = 220;
                p.Text = "個人選單安全驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.White;

                Label lbl = new Label() { Left = 30, Top = 30, Text = $"查看此隱藏選單資料表，請輸入【{menuName}】的解鎖密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 250, Left = 30, Top = 70, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認驗證", DialogResult = DialogResult.OK, Left = 160, Top = 120, Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog() == DialogResult.OK)
                {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName) return true;
                    
                    MessageBox.Sh
