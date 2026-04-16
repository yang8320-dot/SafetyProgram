/// FILE: Safety_System/App_HomeDashboard.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_HomeDashboard
    {
        private FlowLayoutPanel _flp;
        private Panel _main;

        public Control GetView()
        {
            // 1. 主容器：鎖定全滿，開啟自動捲動
            _main = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                AutoScroll = true,
                Padding = new Padding(30)
            };

            // 標題與副標題容器（固定在上方）
            Panel headerPnl = new Panel { Dock = DockStyle.Top, Height = 130, BackColor = Color.Transparent };

            Label lblTitle = new Label {
                Text = "Safety System 綜合數據看板系統",
                Font = new Font("Microsoft JhengHei UI", 28F, FontStyle.Bold),
                ForeColor = Color.FromArgb(45, 62, 80),
                Location = new Point(10, 20),
                AutoSize = true
            };
            
            Label lblSubTitle = new Label {
                Text = "🚀 點選下方快捷入口，即時查閱各項工安與環境管理指標。",
                Font = new Font("Microsoft JhengHei UI", 14F),
                ForeColor = Color.Gray,
                Location = new Point(15, 80),
                AutoSize = true
            };
            headerPnl.Controls.Add(lblTitle);
            headerPnl.Controls.Add(lblSubTitle);

            // 2. 快捷按鈕容器：流式排版 (FlowLayoutPanel)
            _flp = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,      
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                WrapContents = true,        
                Padding = new Padding(0, 20, 0, 50)
            };

            // 監聽主容器 Resize 事件，讓排列面板的寬度永遠跟隨視窗
            _main.Resize += (s, e) => {
                _flp.Width = _main.ClientSize.Width - _main.Padding.Left - _main.Padding.Right;
            };

            // ========================================================
            // 註冊 12 個看板入口 (新增 ISO14001, 刪除操作說明)
            // ========================================================
            _flp.Controls.Add(CreateShortcut("🛡️ 工安看板", "零災害天數、虛驚、工傷統計摘要", Color.SteelBlue, () => new App_SafetyDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("🧪 化學品看板", "SDS清冊、化學品管制與風險警示", Color.DarkCyan, () => new App_ChemDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("🏥 護理看板", "職場健康、活動促進與職災申報統計", Color.PaleVioletRed, () => new App_NursingDashboard().GetView()));
            
            // 🟢 將空污看板顏色調亮為 DeepSkyBlue
            _flp.Controls.Add(CreateShortcut("☁️ 空污看板", "空氣污染防治監測與排放數據總覽", Color.DeepSkyBlue, () => new App_AirDashboard().GetView()));
            
            _flp.Controls.Add(CreateShortcut("💧 水資源看板", "水情摘要、用水量與YoY差異分析", Color.Teal, () => new App_WaterDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("♻️ 廢棄物看板", "產能月表、廢棄物清運與減量趨勢", Color.SeaGreen, () => new App_WasteDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("🔥 消防看板", "消防設備巡檢狀態與火源責任人管理", Color.IndianRed, () => new App_FireDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("📋 檢測看板", "環境、水質、TCLP等各類檢測數據", Color.Chocolate, () => new App_TestDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("🎓 訓練看板", "訓練時數統計、證照管理與績效評估", Color.Goldenrod, () => new App_EduDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("⚖️ 法規看板", "法規鑑別統計、更新進度與分析", Color.SlateGray, () => new App_LawDashboard().GetView()));
            _flp.Controls.Add(CreateShortcut("🌱 ESG看板", "永續發展績效、能源節約與碳排管理", Color.DarkOliveGreen, () => new App_ESGDashboard().GetView()));
            
            // 🟢 新增 ISO14001，取代原本的「操作說明」
            _flp.Controls.Add(CreateShortcut("📐 ISO看板", "環境管理目標、稽核計畫與績效指標", Color.MidnightBlue, () => new App_ISODashboard().GetView()));

            _main.Controls.Add(_flp);
            _main.Controls.Add(headerPnl);

            return _main;
        }

        /// <summary>
        /// 建立具備動態懸停效果的快捷入口卡片
        /// </summary>
        private Control CreateShortcut(string title, string desc, Color themeColor, Func<Control> loadFunc)
        {
            Panel p = new Panel { 
                Size = new Size(320, 170), 
                Margin = new Padding(0, 0, 25, 25), 
                BackColor = Color.FromArgb(250, 250, 252), 
                Cursor = Cursors.Hand 
            };
            
            // 繪製卡片陰影與邊框感
            p.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.FromArgb(230, 230, 230), ButtonBorderStyle.Solid);
                // 頂部主題條
                using (SolidBrush b = new SolidBrush(themeColor)) {
                    e.Graphics.FillRectangle(b, 0, 0, p.Width, 6);
                }
            };

            Label l1 = new Label { 
                Text = title, 
                Font = new Font("Microsoft JhengHei UI", 17F, FontStyle.Bold), 
                ForeColor = themeColor, 
                Location = new Point(15, 25), 
                AutoSize = true 
            };
            
            Label l2 = new Label { 
                Text = desc, 
                Font = new Font("Microsoft JhengHei UI", 11F), 
                ForeColor = Color.DimGray, 
                Location = new Point(17, 70), 
                Size = new Size(280, 50) 
            };
            
            Button btn = new Button {
                Text = "進入看板 ➜",
                Size = new Size(110, 35),
                Location = new Point(190, 120),
                FlatStyle = FlatStyle.Flat,
                BackColor = themeColor,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;

            // 點擊事件統整
            Action onClick = () => {
                Form f = p.FindForm();
                if (f is MainForm mf) mf.LoadModule(loadFunc());
            };

            // 註冊所有子元件點擊
            p.Click += (s, e) => onClick();
            l1.Click += (s, e) => onClick();
            l2.Click += (s, e) => onClick();
            btn.Click += (s, e) => onClick();

            // 簡單的懸停效果
            p.MouseEnter += (s, e) => { p.BackColor = Color.White; p.Top -= 2; };
            p.MouseLeave += (s, e) => { p.BackColor = Color.FromArgb(250, 250, 252); p.Top += 2; };

            p.Controls.Add(l1); 
            p.Controls.Add(l2); 
            p.Controls.Add(btn);
            
            return p;
        }
    }
}
