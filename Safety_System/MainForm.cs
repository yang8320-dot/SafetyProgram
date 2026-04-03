using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "工安系統看板 (v4.7.2 - SQLite 版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30),
                AutoScroll = true
            };
            this.Controls.Add(_contentPanel);
            
            _mainMenu.BringToFront();
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // 1. 首頁
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            // 2. 報表
            var menuReport = new ToolStripMenuItem("報表");
            menuReport.DropDownItems.Add(CreatePlaceholderItem("月報表", "App_MonthlyReport.cs"));
            menuReport.DropDownItems.Add(CreatePlaceholderItem("年報表", "App_YearlyReport.cs"));

            // 3. 工安
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("虛驚", "App_NearMiss.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("工傷", "App_WorkInjury.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("交傷", "App_TrafficInjury.cs"));

            // 4. 環保
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(CreatePlaceholderItem("空污申報", "App_AirPollution.cs"));

            // 5. 消防
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreatePlaceholderItem("消防設備", "App_FireEquip.cs"));

            // 6. 設定
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDb = new ToolStripMenuItem("資料庫設定");
            itemDb.Click += (s, e) => { new App_DbConfig().Show(); };
            menuSettings.DropDownItems.Add(itemDb);

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings
            });
        }

        // 建立預留項目的輔助方法
        private ToolStripMenuItem CreatePlaceholderItem(string text, string fileName)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => ShowPlaceholder(fileName);
            return item;
        }

        private void ShowPlaceholder(string fileName)
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label
            {
                Text = string.Format("本頁面對應 {0}, 資料尚在建立中", fileName),
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Italic),
                ForeColor = Color.DarkGray,
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }

        public void LoadModule(Control moduleControl)
        {
            _contentPanel.Controls.Clear();
            moduleControl.Dock = DockStyle.Fill;
            _contentPanel.Controls.Add(moduleControl);
        }

        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n解析度：1440 x 810",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
