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
            this.Text = "工安系統看板 (v4.8.6 - 模組增修版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(20),
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

            // --- 3. 工安 (更新後的選單排序) ---
            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安儀表版", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄", () => new App_SafetyInspection().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件", () => new App_NearMiss().GetView()));
            // 🟢 新增：安全觀查
            menuSafety.DropDownItems.Add(CreateItem("安全觀查", () => new App_SafetyObservation().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件", () => new App_WorkInjury().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交傷事件", () => new App_TrafficInjury().GetView()));

            // --- 4. 護理 ---
            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理儀表版", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報數據", () => new App_WorkInjuryReport().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進", () => new App_HealthPromotion().GetView()));

            // --- 5. 空污 ---
            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污儀表版", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報", () => new App_AirPollution().GetView()));

            // --- 6. 水 ---
            var menuWater = new ToolStripMenuItem("水");
            menuWater.DropDownItems.Add(CreateItem("水資源儀表版", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("水處理記錄表", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("用藥記錄表", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("自水水量", () => new App_WaterVolume().GetView()));
            menuWater.DropDownItems.Add(CreateItem("納管排放數據", () => new App_DischargeData().GetView()));

            // --- 7. 廢棄物 ---
            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物儀表版", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("生產月報表", () => new App_WasteMonthly().GetView()));

            // --- 8. 消防 ---
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防儀表版", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備", () => new App_FireEquip().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人", () => new App_FireResponsible().GetView()));
            menuFire.DropDownItems.Add(CreateItem("公危統計表", () => new App_HazardStats().GetView()));

            // --- 9. 設定 ---
            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("說明", () => new App_Instruction().GetView()));

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuNursing, menuAir, menuWater, menuWaste, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings
            });
        }

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
            Label lbl = new Label {
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n(已更新：工安選單納入安全觀查功能)",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, 
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
