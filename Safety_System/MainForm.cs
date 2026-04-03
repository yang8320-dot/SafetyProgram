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
            this.Text = "工安系統看板 (v4.8.0 - 結構優化版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            _contentPanel = new Panel {
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
            // --- 1. 首頁 ---
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            // --- 2. 報表 ---
            var menuReport = new ToolStripMenuItem("報表");
            menuReport.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReport.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            // --- 3. 工安 ---
            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("巡檢", () => new App_SafetyInspection().GetView()));

            // --- 4. 護理 (新增於工安旁) ---
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

            // --- 7. 廢棄物 (新增於水旁) ---
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

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuNursing, menuAir, menuWater, menuWaste, menuFire, menuSettings
            });
        }

        // 輔助方法：簡化選單建立邏輯
        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => LoadModule(getViewFunc());
            return item;
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
            Label lbl = new Label { Text = "工安系統看板 v4.8\n歡迎使用", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(50, 50) };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
