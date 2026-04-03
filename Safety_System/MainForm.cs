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
            // 🟢 注意：全域快捷鍵功能已移除，由 Program.cs 處理重複啟動時的自動喚醒
        }

        private void InitializeComponent()
        {
            this.Text = "工安系統看板 (v4.7.2 - SQLite 版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            // 初始化選單欄
            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 初始化動態內容容器
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30),
                AutoScroll = true
            };
            this.Controls.Add(_contentPanel);
            
            _mainMenu.BringToFront();

            // 預設載入歡迎畫面
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // --- 1. 首頁 ---
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            // --- 2. 報表 ---
            var menuReport = new ToolStripMenuItem("報表");
            var itemMonthly = new ToolStripMenuItem("月報表");
            itemMonthly.Click += (s, e) => LoadModule(new App_MonthlyReport().GetView());
            var itemYearly = new ToolStripMenuItem("年報表");
            itemYearly.Click += (s, e) => LoadModule(new App_YearlyReport().GetView());
            menuReport.DropDownItems.AddRange(new ToolStripItem[] { itemMonthly, itemYearly });

            // --- 3. 工安 ---
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("虛驚", "App_NearMiss.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("工傷", "App_WorkInjury.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("交傷", "App_TrafficInjury.cs"));

            // --- 4. 環保 ---
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(CreatePlaceholderItem("空污申報", "App_AirPollution.cs"));

            // --- 5. 水資源 (排序優化) ---
            var menuWater = new ToolStripMenuItem("水資源");
            
            var itemWaterDashboard = new ToolStripMenuItem("水資源儀表版");
            itemWaterDashboard.Click += (s, e) => LoadModule(new App_WaterDashboard().GetView());

            var itemWaterTreat = new ToolStripMenuItem("水處理記錄表");
            itemWaterTreat.Click += (s, e) => LoadModule(new App_WaterTreatment().GetView());
            
            var itemWaterChem = new ToolStripMenuItem("用藥記錄表");
            itemWaterChem.Click += (s, e) => LoadModule(new App_WaterChemicals().GetView());
            
            var itemWaterVol = new ToolStripMenuItem("自水水量");
            itemWaterVol.Click += (s, e) => LoadModule(new App_WaterVolume().GetView());

            var itemDischarge = new ToolStripMenuItem("納管排放數據");
            itemDischarge.Click += (s, e) => LoadModule(new App_DischargeData().GetView());
            
            // 確保儀表版排在第一個
            menuWater.DropDownItems.AddRange(new ToolStripItem[] { 
                itemWaterDashboard, 
                itemWaterTreat, 
                itemWaterChem, 
                itemWaterVol, 
                itemDischarge 
            });

            // --- 6. 消防 ---
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreatePlaceholderItem("消防設備", "App_FireEquip.cs"));

            // --- 7. 設定 ---
            var menuSettings = new ToolStripMenuItem("設定");
            
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            itemDbConfig.Click += (s, e) => { new App_DbConfig().Show(); };
            
            var itemInstruction = new ToolStripMenuItem("說明");
            itemInstruction.Click += (s, e) => LoadModule(new App_Instruction().GetView());

            menuSettings.DropDownItems.AddRange(new ToolStripItem[] { itemDbConfig, itemInstruction });

            // 將所有主選單加入選單列
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuEnv, menuWater, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings
            });
        }

        /// <summary>
        /// 動態切換顯示的模組視圖
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
        /// 載入首頁歡迎畫面
        /// </summary>
        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label {
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n(偵測重複開啟：自動喚醒舊視窗模式)",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, 
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }

        /// <summary>
        /// 建立尚在開發中模組的佔位符
        /// </summary>
        private ToolStripMenuItem CreatePlaceholderItem(string text, string fileName)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                _contentPanel.Controls.Clear();
                Label lbl = new Label {
                    Text = string.Format("本頁面對應 {0}, 資料尚在建立中", fileName),
                    Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Italic),
                    ForeColor = Color.DarkGray, AutoSize = true, Location = new Point(50, 50)
                };
                _contentPanel.Controls.Add(lbl);
            };
            return item;
        }
    }
}
