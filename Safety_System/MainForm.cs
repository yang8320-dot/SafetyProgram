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
            this.Text = "工安系統看板 (v6.0 - 全面 Generic 共用模組化)";
            
            // 設定初始視窗為最大化
            this.WindowState = FormWindowState.Maximized;
            
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);
            
            DataManager.LoadConfig();
            
            // 🟢 啟動時自動檢查是否需要執行備份
            BackupManager.RunAutoBackup();

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

        // 🟢 支援全局 Ctrl+S 快捷存檔
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

        // 🟢 建立主選單 (已全面導入 App_GenericTable)
        private void BuildMenu()
        {
            var menuHome = new ToolStripMenuItem("頁首");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuReports = new ToolStripMenuItem("報表");
            menuReports.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReports.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            // 🟢 工安選單名稱修正與重新排序，並新增「輕傷事件」
            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄", () => new App_GenericTable("Safety", "SafetyInspection", "巡檢記錄管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件", () => new App_GenericTable("Safety", "WorkInjury", "工傷事件管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("輕傷事件", () => new App_GenericTable("Safety", "MinorInjury", "輕傷事件管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交通意外", () => new App_GenericTable("Safety", "TrafficInjury", "交通意外紀錄").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件", () => new App_GenericTable("Safety", "NearMiss", "虛驚事件管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("安全觀察", () => new App_GenericTable("Safety", "SafetyObservation", "安全觀察紀錄").GetView()));

            var menuChemical = new ToolStripMenuItem("化學品");
            menuChemical.DropDownItems.Add(CreateItem("化學品看板", () => new App_ChemDashboard().GetView()));
            menuChemical.DropDownItems.Add(CreateItem("化學品快查", () => new App_ChemQuickSearch().GetView()));
            menuChemical.DropDownItems.Add(CreateItem("化學品要求及規範", () => new App_GenericTable("Chemical", "ChemRegulations", "化學品要求及規範").GetView()));
            menuChemical.DropDownItems.Add(CreateItem("SDS清冊", () => new App_GenericTable("Chemical", "SDS_Inventory", "SDS清冊").GetView()));

            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進活動", () => new App_GenericTable("Nursing", "HealthPromotion", "健康促進活動").GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報紀錄", () => new App_GenericTable("Nursing", "WorkInjuryReport", "職災申報紀錄").GetView()));

            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報紀錄", () => new App_GenericTable("Air", "AirPollution", "空污申報紀錄").GetView()));

            var menuWater = new ToolStripMenuItem("水污");
            menuWater.DropDownItems.Add(CreateItem("水資源管理看板", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理水量記錄", () => new App_Water_Generic("Water", "WaterMeterReadings", "【日】廢水處理水量記錄").GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理用藥記錄", () => new App_Water_Generic("Water", "WaterChemicals", "【日】廢水處理用藥記錄").GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】自來水使用量", () => new App_Water_Generic("Water", "WaterUsageDaily", "【日】自來水使用量").GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】納管排放數據", () => new App_Water_Generic("Water", "DischargeData", "【月】納管排放數據").GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】自來水用量統計", () => new App_Water_Generic("Water", "WaterVolume", "【月】自來水用量統計").GetView()));

            var menuWaste = new ToolStripMenuItem("產能及廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("產能及廢棄物看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("複層月表", () => new App_GenericTable("Waste", "Waste_IL", "複層月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("膠合月表", () => new App_GenericTable("Waste", "Waste_LM", "膠合月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("鍍板月表", () => new App_GenericTable("Waste", "Waste_CR", "鍍板月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("強化月表", () => new App_GenericTable("Waste", "Waste_T", "強化月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("切磨月表", () => new App_GenericTable("Waste", "Waste_GCTE", "切磨月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("物料月表", () => new App_GenericTable("Waste", "Waste_ML", "物料月表").GetView()));
            menuWaste.DropDownItems.Add(CreateItem("水站月表", () => new App_GenericTable("Waste", "Waste_Water", "水站月表").GetView()));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人管理", () => new App_GenericTable("Fire", "FireResponsible", "火源責任人管理").GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_GenericTable("Fire", "HazardStats", "公共危險物統計").GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_GenericTable("Fire", "FireEquip", "消防設備巡檢").GetView()));

            var menuTest = new ToolStripMenuItem("檢測數據");
            menuTest.DropDownItems.Add(CreateItem("檢測數據看版", () => new App_TestDashboard().GetView()));
            menuTest.DropDownItems.Add(CreateItem("環境監測", () => new App_GenericTable("TestData", "EnvMonitor", "環境監測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水定申檢", () => new App_GenericTable("TestData", "WastewaterPeriodic", "廢水定申檢").GetView()));
            menuTest.DropDownItems.Add(CreateItem("飲用水檢測", () => new App_GenericTable("TestData", "DrinkingWater", "飲用水檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("工業區檢驗", () => new App_GenericTable("TestData", "IndustrialZoneTest", "工業區檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("土壤氣體檢測", () => new App_GenericTable("TestData", "SoilGasTest", "土壤氣體檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水自主檢驗", () => new App_GenericTable("TestData", "WastewaterSelfTest", "廢水自主檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(廠商)", () => new App_GenericTable("TestData", "CoolingWaterVendor", "循環水檢測(廠商)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(自評)", () => new App_GenericTable("TestData", "CoolingWaterSelf", "循環水檢測(自評)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("TCLP", () => new App_GenericTable("TestData", "TCLP", "TCLP毒性特性溶出").GetView()));
            menuTest.DropDownItems.Add(CreateItem("水錶校正", () => new App_GenericTable("TestData", "WaterMeterCalibration", "水錶校正").GetView()));
            menuTest.DropDownItems.Add(CreateItem("其它檢測數據", () => new App_GenericTable("TestData", "OtherTests", "其它檢測數據").GetView()));

            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(CreateItem("訓練時數", () => new App_GenericTable("教育訓練", "訓練時數", "教育訓練時數").GetView()));

            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("法規看板", () => new App_LawDashboard().GetView()));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "環保法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "職安衛法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "其它法規"));

            var menuESG = new ToolStripMenuItem("ESG");
            menuESG.DropDownItems.Add(CreateItem("ESG看板", () => new App_ESGDashboard().GetView())); 
            menuESG.DropDownItems.Add(CreateItem("ESG績效管理", () => new App_GenericTable("ESG", "ESG_Performance", "ESG績效管理").GetView()));

            // 🟢 新增：ISO14001 選單 (位於 ESG 後)
            var menuISO = new ToolStripMenuItem("ISO14001");
            menuISO.DropDownItems.Add(CreateItem("ISO看板", () => new App_ISODashboard().GetView()));
            menuISO.DropDownItems.Add(CreateItem("目標管理", () => new App_GenericTable("ISO14001", "TargetManagement", "目標管理").GetView()));

            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    // 呼叫全域的 AuthManager，要求管理者權限
                    if (AuthManager.VerifyAdmin()) { LoadModule(new App_DbConfig().GetView()); } 
                    else { MessageBox.Show("密碼錯誤或權限不足，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            // 🟢 將選單加入主視窗
            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuChemical, menuNursing, menuAir, 
                menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuESG, menuISO, menuSettings 
            });
        }

        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                try { LoadModule(getViewFunc()); } 
                catch (Exception ex) { MessageBox.Show($"載入模組 {text} 失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
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
