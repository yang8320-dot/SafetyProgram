/// FILE: Safety_System/MainForm.cs ///
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
            this.Text = "工安系統看板 (v5.0.8 - 驗證視窗優化版)";
            
            // 🟢 設定初始視窗為最大化
            this.WindowState = FormWindowState.Maximized;
            
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);
            DataManager.LoadConfig();

            _mainMenu = new MenuStrip { Font = new Font("Microsoft JhengHei UI", 12F), Dock = DockStyle.Top };
            BuildMenu();

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(15, 15, 15, 15),
                AutoScroll = true
            };

            this.Controls.Add(_contentPanel);
            this.Controls.Add(_mainMenu); 
            
            LoadWelcomeScreen();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                Button activeSaveButton = FindControlByName(_contentPanel, "btnSave") as Button;
                if (activeSaveButton != null && activeSaveButton.Enabled)
                {
                    this.Validate(); 
                    activeSaveButton.PerformClick();
                    return true; 
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private Control FindControlByName(Control parent, string name)
        {
            foreach (Control c in parent.Controls)
            {
                if (c.Name == name) return c;
                Control found = FindControlByName(c, name);
                if (found != null) return found;
            }
            return null;
        }

        private void BuildMenu()
        {
            var menuHome = new ToolStripMenuItem("頁首");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuReports = new ToolStripMenuItem("報表");
            menuReports.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReports.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件管理", () => new App_NearMiss().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄管理", () => new App_SafetyInspection().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("安全觀察紀錄", () => new App_SafetyObservation().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交通意外紀錄", () => new App_TrafficInjury().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件管理", () => new App_WorkInjury().GetView()));

            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進活動", () => new App_HealthPromotion().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報紀錄", () => new App_WorkInjuryReport().GetView()));

            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報紀錄", () => new App_AirPollution().GetView()));

            var menuWater = new ToolStripMenuItem("水污");
            menuWater.DropDownItems.Add(CreateItem("水資源管理看板", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理水量記錄", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理用藥記錄", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】自來水使用量", () => new App_WaterUsageDaily().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】納管排放數據", () => new App_DischargeData().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】自來水用量統計", () => new App_WaterVolume().GetView()));

            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("廢棄物統計表", () => new App_WasteMonthly().GetView()));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人管理", () => new App_FireResponsible().GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_HazardStats().GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_FireEquip().GetView()));

            var menuTest = new ToolStripMenuItem("檢測數據");
            menuTest.DropDownItems.Add(CreateItem("檢測數據看版", () => new App_TestDashboard().GetView()));
            
            // 🟢 全面改用 Generic 共用模組，傳入: 資料庫名, 表名, 中文標題
            menuTest.DropDownItems.Add(CreateItem("環境監測", () => new App_Test_Generic("TestData", "EnvMonitor", "環境監測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水定申檢", () => new App_Test_Generic("TestData", "WastewaterPeriodic", "廢水定申檢").GetView()));
            menuTest.DropDownItems.Add(CreateItem("飲用水檢測", () => new App_Test_Generic("TestData", "DrinkingWater", "飲用水檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("工業區檢驗", () => new App_Test_Generic("TestData", "IndustrialZoneTest", "工業區檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("土壤氣體檢測", () => new App_Test_Generic("TestData", "SoilGasTest", "土壤氣體檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水自主檢驗", () => new App_Test_Generic("TestData", "WastewaterSelfTest", "廢水自主檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(廠商)", () => new App_Test_Generic("TestData", "CoolingWaterVendor", "循環水檢測(廠商)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(自評)", () => new App_Test_Generic("TestData", "CoolingWaterSelf", "循環水檢測(自評)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("TCLP", () => new App_Test_Generic("TestData", "TCLP", "TCLP毒性特性溶出").GetView()));
            menuTest.DropDownItems.Add(CreateItem("水錶校正", () => new App_Test_Generic("TestData", "WaterMeterCalibration", "水錶校正").GetView()));
            menuTest.DropDownItems.Add(CreateItem("其它檢測數據", () => new App_Test_Generic("TestData", "OtherTests", "其它檢測數據").GetView()));

            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(CreateItem("訓練時數", () => new App_EduHours().GetView()));

            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("法規看板", () => new App_LawDashboard().GetView()));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "環保法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "職安衛法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "其它法規"));

            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    // 🟢 呼叫全域的 AuthManager，要求管理者權限 (11914002)
                    if (AuthManager.VerifyAdmin()) { LoadModule(new App_DbConfig().GetView()); } 
                    else { MessageBox.Show("密碼錯誤或權限不足，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuNursing, menuAir, 
                menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuSettings 
            });
        }

        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                try { LoadModule(getViewFunc()); } 
                catch (Exception ex) { MessageBox.Show($"載入模組 {text} 失敗：\n{ex.Message}\n(您可能尚未建立此功能的 CS 檔案)", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            return item;
        }

        private ToolStripMenuItem CreateLawItem(string dbName, string tableName)
        {
            var item = new ToolStripMenuItem(tableName);
            item.Click += (s, e) => {
                try { LoadModule(new App_Law_Generic(dbName, tableName).GetView()); } 
                catch (Exception ex) { MessageBox.Show($"載入模組失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            return item;
        }

        public void LoadModule(Control moduleControl)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => LoadModule(moduleControl))); return; }
            if (moduleControl == null) return;
            try {
                _contentPanel.Controls.Clear();
                moduleControl.Dock = DockStyle.Fill;
                _contentPanel.Controls.Add(moduleControl);
            } catch (Exception ex) {
                MessageBox.Show($"畫面切換時發生錯誤：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadWelcomeScreen()
        {
            try { LoadModule(new App_HomeDashboard().GetView()); } 
            catch { }
        }
    }
}
