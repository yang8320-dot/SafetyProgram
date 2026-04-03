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
            this.Text = "工安系統看板 (v4.7.2)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);
            this.Font = new Font("Microsoft JhengHei UI", 12F);
            this.KeyPreview = true;

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

            // 🟢 2. 報表 (新插入：包含月報表、年報表)
            var menuReport = new ToolStripMenuItem("報表");
            var itemMonthly = new ToolStripMenuItem("月報表");
            var itemYearly = new ToolStripMenuItem("年報表");
            
            // 綁定點擊事件
            itemMonthly.Click += (s, e) => LoadModule(new App_Report().GetMonthlyReportView());
            itemYearly.Click += (s, e) => LoadModule(new App_Report().GetYearlyReportView());
            
            menuReport.DropDownItems.Add(itemMonthly);
            menuReport.DropDownItems.Add(itemYearly);

            // 3. 工安
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("虛驚"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("工傷"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("交傷"));

            // 4. 其他選單 (環保、消防、設定...)
            var menuEnv = new ToolStripMenuItem("環保");
            var menuFire = new ToolStripMenuItem("消防");
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            itemDbConfig.Click += (s, e) => { new App_DbConfig().Show(); };
            menuSettings.DropDownItems.Add(itemDbConfig);

            // 將選單按順序加入
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, 
                menuReport, // 報表放在工安前
                menuSafety, 
                menuEnv, 
                menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                menuSettings
            });
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.D3) || keyData == (Keys.Control | Keys.NumPad3))
            {
                LoadModule(new App_SafetyInspection().GetView());
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        public void LoadModule(Control moduleControl)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => LoadModule(moduleControl)));
                return;
            }
            _contentPanel.Controls.Clear();
            moduleControl.Dock = DockStyle.Fill;
            _contentPanel.Controls.Add(moduleControl);
        }

        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n已進入報表管理架構\n(快捷鍵: Ctrl+3 開啟巡檢)",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
