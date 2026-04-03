using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 主視窗：負責整體佈局與模組畫面切換 (已移除 Ctrl + 3 功能)
    /// </summary>
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
            // 視窗基本設定為 1440 x 810
            this.Text = "工安系統看板 (v4.7.2)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            // 設定系統預設字體以確保清晰度
            this.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));

            // 初始化選單
            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 主顯示容器 (大框)
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
            var itemMonthly = new ToolStripMenuItem("月報表");
            var itemYearly = new ToolStripMenuItem("年報表");
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

            // 4. 環境與消防
            var menuEnv = new ToolStripMenuItem("環保");
            var menuFire = new ToolStripMenuItem("消防");

            // 5. 設定
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            itemDbConfig.Click += (s, e) => { new App_DbConfig().Show(); };
            menuSettings.DropDownItems.Add(itemDbConfig);

            // 將所有選單加入列
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), 
                menuSettings
            });
        }

        /// <summary>
        /// 🟢 已修正：移除 ProcessCmdKey 中的 Ctrl + 3 判斷邏輯
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 目前無特殊快捷鍵需求，直接回傳基底處理
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

        /// <summary>
        /// 🟢 已修正：移除主頁文字中的快捷鍵提示
        /// </summary>
        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n架構：.NET Framework 4.7.2\n系統環境已就緒，請由上方選單開始作業。",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
