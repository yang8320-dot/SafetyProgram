using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 主視窗：系統核心導覽外殼
    /// 支援解析度：1440 x 810
    /// 資料引擎：SQLite
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
            // 1. 視窗基本設定
            this.Text = "工安系統看板 (v4.7.2 - SQLite 版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            // 設定全系統 UI 清晰字體
            this.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));

            // 2. 初始化選單列
            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 3. 初始化主內容容器 (所有功能畫面都載入到這裡)
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30),
                AutoScroll = true
            };
            this.Controls.Add(_contentPanel);
            
            // 確保選單始終在最上層
            _mainMenu.BringToFront();

            // 4. 載入初始歡迎畫面
            LoadWelcomeScreen();
        }

        /// <summary>
        /// 建構選單結構
        /// </summary>
        private void BuildMenu()
        {
            // --- 選單 1：首頁 (直接按鍵) ---
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            // --- 選單 2：報表 (下拉選單) ---
            var menuReport = new ToolStripMenuItem("報表");
            var itemMonthly = new ToolStripMenuItem("月報表");
            itemMonthly.Click += (s, e) => LoadModule(new App_MonthlyReport().GetView());
            
            var itemYearly = new ToolStripMenuItem("年報表");
            itemYearly.Click += (s, e) => LoadModule(new App_YearlyReport().GetView());
            
            menuReport.DropDownItems.AddRange(new ToolStripItem[] { itemMonthly, itemYearly });

            // --- 選單 3：工安 (下拉選單) ---
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("虛驚", "App_NearMiss.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("工傷", "App_WorkInjury.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("交傷", "App_TrafficInjury.cs"));

            // --- 選單 4：環保 ---
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(CreatePlaceholderItem("空污申報", "App_AirPollution.cs"));

            // --- 選單 5：消防 ---
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreatePlaceholderItem("消防設備", "App_FireEquip.cs"));

            // --- 選單 6：設定 (包含資料庫路徑設定) ---
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            // 使用 .Show() 確保視窗切換不卡住
            itemDbConfig.Click += (s, e) => { 
                App_DbConfig configForm = new App_DbConfig();
                configForm.Show(); 
            };
            menuSettings.DropDownItems.Add(itemDbConfig);

            // 將所有選單依序加入 MenuStrip
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, 
                menuReport,
                menuSafety, 
                menuEnv, 
                menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                menuSettings
            });
        }

        /// <summary>
        /// 輔助方法：建立尚未開發完成的預留選項
        /// </summary>
        private ToolStripMenuItem CreatePlaceholderItem(string text, string fileName)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
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
            };
            return item;
        }

        /// <summary>
        /// 動態切換功能模組畫面
        /// </summary>
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
        /// 回到初始歡迎頁面
        /// </summary>
        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n\n資料引擎：SQLite (已連結)\n系統環境已就緒，請由上方選單開始作業。",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }

        /// <summary>
        /// 處理快捷鍵 (目前已移除 Ctrl+3 邏輯)
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
