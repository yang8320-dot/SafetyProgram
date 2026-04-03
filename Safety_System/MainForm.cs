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
            // 🟢 注意：全域快捷鍵 Ctrl+3 已移除，重複啟動喚醒邏輯已移至 Program.cs
        }

        private void InitializeComponent()
        {
            this.Text = "工安系統看板 (v4.8.5 - 結構最優化版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            // 1. 初始化選單欄
            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 2. 初始化主內容容器
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(20),
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
            menuReport.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReport.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            // --- 3. 工安 (重新修訂) ---
            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安儀表版", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄", () => new App_SafetyInspection().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件", () => new App_NearMiss().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件", () => new App_WorkInjury().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交傷事件", () => new App_TrafficInjury().GetView()));

            // --- 4. 護理 (新增) ---
            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理儀表版", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報數據", () => new App_WorkInjuryReport().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進", () => new App_HealthPromotion().GetView()));

            // --- 5. 空污 (原環保改名) ---
            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污儀表版", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報", () => new App_AirPollution().GetView()));

            // --- 6. 水 (原水資源改名) ---
            var menuWater = new ToolStripMenuItem("水");
            menuWater.DropDownItems.Add(CreateItem("水資源儀表版", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("水處理記錄表", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("用藥記錄表", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("自水水量", () => new App_WaterVolume().GetView()));
            menuWater.DropDownItems.Add(CreateItem("納管排放數據", () => new App_DischargeData().GetView()));

            // --- 7. 廢棄物 (新增) ---
            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物儀表版", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("生產月報表", () => new App_WasteMonthly().GetView()));

            // --- 8. 消防 (重新排序與新增) ---
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防儀表版", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備", () => new App_FireEquip().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人", () => new App_FireResponsible().GetView()));
            menuFire.DropDownItems.Add(CreateItem("公危統計表", () => new App_HazardStats().GetView()));

            // --- 9. 設定 ---
            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("說明", () => new App_Instruction().GetView()));

            // 將所有選單加入 MenuStrip
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuNursing, menuAir, menuWater, menuWaste, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings
            });
        }

        /// <summary>
        /// 輔助方法：快速建立選單項目並連結對應的 .cs 類別
        /// </summary>
        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                try {
                    LoadModule(getViewFunc());
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入模組 {text}：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            return item;
        }

        /// <summary>
        /// 切換內容面板顯示的模組
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
        /// 顯示首頁畫面
        /// </summary>
        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label {
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n系統狀態：已啟動並偵測單一實例",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, 
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
