/// FILE: Safety_System/MainForm.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        private ToolStripMenuItem _menu1;
        private ToolStripMenuItem _menu2;
        private ToolStripMenuItem _menu3;
        private ToolStripMenuItem _menu4;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "工安系統看板";
            
            this.WindowState = FormWindowState.Maximized;
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);
            
            DataManager.LoadConfig();

            App_DropdownManager.LoadDropdownConfigs();
            App_DropdownManager.LoadMultiSelectConfigs();
            
            Task.Run(() => {
                try { BackupManager.RunAutoBackup(); } catch { }
            });
            
            App_PasswordManager.InitDatabase();

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
            
            this.Shown += MainForm_Shown;
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            LoadWelcomeScreen();
            
            Task.Delay(500).ContinueWith(_ => {
                try {
                    ReminderEngine.CheckAndShowReminders();
                } catch { }
            });
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            ForceEndCurrentEdit();
            base.OnFormClosing(e);
            Environment.Exit(0);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                Button activeSaveButton = FindControlByName(_contentPanel, "btnSave") as Button;
                if (activeSaveButton != null && activeSaveButton.Enabled)
                {
                    ForceEndCurrentEdit();
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

        private void ForceEndCurrentEdit()
        {
            this.Validate();
            
            void SearchAndEndEdit(Control parent)
            {
                foreach (Control c in parent.Controls)
                {
                    if (c is DataGridView dgv && dgv.IsCurrentCellInEditMode)
                    {
                        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                        dgv.EndEdit();
                    }
                    else if (c.HasChildren)
                    {
                        SearchAndEndEdit(c);
                    }
                }
            }
            
            SearchAndEndEdit(_contentPanel);
        }

        private void BuildMenu()
        {
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuReports = new ToolStripMenuItem("日常作業");
            menuReports.DropDownItems.Add(CreateItem("A11 請購資料", () => new App_CoreTable("Purchase", "PurchaseData", "請購資料", new DefaultLogic()).GetView())); 
            menuReports.DropDownItems.Add(CreateItem("A12 月報表", () => new App_MonthlyReport().GetView()));
            menuReports.DropDownItems.Add(CreateItem("A13 年報表", () => new App_YearlyReport().GetView()));

            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("B11 工安看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("B12 稽核資料查詢", () => new App_AuditDashboard().GetView()));

            menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("B21 巡檢記錄", () => new App_CoreTable("Safety", "SafetyInspection", "巡檢記錄管理", new DefaultLogic()).GetView()));
			menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("B22 工傷事件", () => new App_CoreTable("Safety", "WorkInjury", "工傷事件管理", new DefaultLogic()).GetView()));
			menuSafety.DropDownItems.Add(CreateItem("B23 交通意外", () => new App_CoreTable("Safety", "TrafficInjury", "交通意外紀錄", new DefaultLogic()).GetView()));
			menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("B24 輕傷事件", () => new App_CoreTable("Safety", "MinorInjury", "輕傷事件管理", new DefaultLogic()).GetView()));  
            menuSafety.DropDownItems.Add(CreateItem("B25 虛驚事件", () => new App_CoreTable("Safety", "NearMiss", "虛驚事件管理", new DefaultLogic()).GetView()));
			menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("B26 安全觀察", () => new App_CoreTable("Safety", "SafetyObservation", "安全觀察紀錄", new DefaultLogic()).GetView()));
            
            menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("B31 勞檢稽查缺失", () => new App_CoreTable("Safety", "LaborInspection", "勞檢稽查缺失", new DefaultLogic()).GetView()));

            var menuChemical = new ToolStripMenuItem("化學品");
            menuChemical.DropDownItems.Add(CreateItem("C11 化學品看板", () => new App_ChemDashboard().GetView()));
            menuChemical.DropDownItems.Add(CreateItem("C12 化學品快查", () => new App_ChemQuickSearch().GetView()));
            menuChemical.DropDownItems.Add(new ToolStripSeparator());

            var menuChemReg = new ToolStripMenuItem("化學品要求及規範");
            menuChemReg.DropDownItems.Add(CreateItem("C21 環測項目", () => new App_CoreTable("Chemical", "EnvTesting", "環測項目", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C22 勞工暴露容許濃度", () => new App_CoreTable("Chemical", "ExposureLimits", "勞工暴露容許濃度", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C23 毒性物質", () => new App_CoreTable("Chemical", "ToxicSubstances", "毒性物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C24 關注性化學物質", () => new App_CoreTable("Chemical", "ConcernedChem", "關注性化學物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C25 優先管理化學品", () => new App_CoreTable("Chemical", "PriorityMgmtChem", "優先管理化學品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C26 管制化學品", () => new App_CoreTable("Chemical", "ControlledChem", "管制化學品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C27 特定化學物質", () => new App_CoreTable("Chemical", "SpecificChem", "特定化學物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C28 有機溶劑", () => new App_CoreTable("Chemical", "OrganicSolvents", "有機溶劑", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C29 勞工健康保護", () => new App_CoreTable("Chemical", "WorkerHealthProtect", "勞工健康保護", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C30 公共危險物品", () => new App_CoreTable("Chemical", "PublicHazardous", "公共危險物品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C31 空污緊急應變", () => new App_CoreTable("Chemical", "AirPollutionEmerg", "空污緊急應變", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("C32 工廠危險物品申報", () => new App_CoreTable("Chemical", "FactoryHazardous", "工廠危險物品申報", new DefaultLogic()).GetView()));
            menuChemical.DropDownItems.Add(menuChemReg);
            
            menuChemical.DropDownItems.Add(CreateItem("C41 SDS清冊", () => new App_CoreTable("Chemical", "SDS_Inventory", "SDS清冊", new DefaultLogic()).GetView()));

            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("D11 護理看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(new ToolStripSeparator());
            menuNursing.DropDownItems.Add(CreateItem("D21 健康促進活動", () => new App_CoreTable("Nursing", "HealthPromotion", "健康促進活動", new DefaultLogic()).GetView()));
            menuNursing.DropDownItems.Add(CreateItem("D22 職災申報紀錄", () => new App_CoreTable("Nursing", "WorkInjuryReport", "職災申報紀錄", new DefaultLogic()).GetView()));

            var menuAirWaterWaste = new ToolStripMenuItem("空水廢");

            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("E11 空污看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(new ToolStripSeparator());
            menuAir.DropDownItems.Add(CreateItem("E21 空污申報紀錄", () => new App_CoreTable("Air", "AirPollution", "空污申報紀錄", new DefaultLogic()).GetView()));

            var menuWater = new ToolStripMenuItem("水污");
            menuWater.DropDownItems.Add(CreateItem("E31 水資源管理看板", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("E32 用水申報", () => new App_WaterReport().GetView()));
            menuWater.DropDownItems.Add(CreateItem("E33 水資源成本表", () => new App_WaterCost().GetView()));
            menuWater.DropDownItems.Add(new ToolStripSeparator());
            menuWater.DropDownItems.Add(CreateItem("E41【日】廢水處理水量記錄", () => new App_CoreTable("Water", "WaterMeterReadings", "【日】廢水處理水量記錄", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("E42【日】廢水處理用藥記錄", () => new App_CoreTable("Water", "WaterChemicals", "【日】廢水處理用藥記錄", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("E43【日】自來水使用量", () => new App_CoreTable("Water", "WaterUsageDaily", "【日】自來水使用量", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("E44【月】納管排放數據", () => new App_CoreTable("Water", "DischargeData", "【月】納管排放數據", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("E45【月】自來水用量統計", () => new App_CoreTable("Water", "WaterVolume", "【月】自來水用量統計", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(new ToolStripSeparator());
            menuWater.DropDownItems.Add(CreateItem("E91 水污許可(原物料)", () => new App_CoreTable("Water", "WaterPermitMaterial", "水污許可(原物料)", new DefaultLogic()).GetView()));

            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("E51 廢棄物管理看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(new ToolStripSeparator());
            menuWaste.DropDownItems.Add(CreateItem("E61【月】複層月表", () => new App_CoreTable("Waste", "Waste_IL", "【月】複層月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E62【月】膠合月表", () => new App_CoreTable("Waste", "Waste_LM", "【月】膠合月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E63【月】鍍板月表", () => new App_CoreTable("Waste", "Waste_CR", "【月】鍍板月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E64【月】強化月表", () => new App_CoreTable("Waste", "Waste_T", "【月】強化月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E65【月】切磨月表", () => new App_CoreTable("Waste", "Waste_GCTE", "【月】切磨月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E66【月】物料月表", () => new App_CoreTable("Waste", "Waste_ML", "【月】物料月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E67【月】水站月表", () => new App_CoreTable("Waste", "Waste_Water", "【月】水站月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(new ToolStripSeparator());
            menuWaste.DropDownItems.Add(CreateItem("E92 廢棄物污許可(原物料)", () => new App_CoreTable("Waste", "WastePermitMaterial", "廢棄物污許可(原物料)", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E93 廢棄物污許可(產品)", () => new App_CoreTable("Waste", "WastePermitProduct", "廢棄物污許可(產品)", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("E94 廢棄物污許可(廢棄物)", () => new App_CoreTable("Waste", "WastePermitWaste", "廢棄物污許可(廢棄物)", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(new ToolStripSeparator());
            menuWaste.DropDownItems.Add(CreateItem("E95 廢棄物清運記錄", () => new App_CoreTable("Waste", "WasteDisposalRecord", "廢棄物清運記錄", new DefaultLogic()).GetView()));

            menuAirWaterWaste.DropDownItems.Add(menuAir);
            menuAirWaterWaste.DropDownItems.Add(menuWater);
            menuAirWaterWaste.DropDownItems.Add(menuWaste);

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("F11 消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(new ToolStripSeparator());
            menuFire.DropDownItems.Add(CreateItem("F21 火源責任人管理", () => new App_CoreTable("Fire", "FireResponsible", "火源責任人管理", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("F22 公共危險物統計", () => new App_CoreTable("Fire", "HazardStats", "公共危險物統計", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("F23 消防設備巡檢", () => new App_CoreTable("Fire", "FireEquip", "消防設備巡檢", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("F24 各單位消防自主檢查表", () => new App_CoreTable("Fire", "FireSelfInspection", "各單位消防自主檢查表", new DefaultLogic()).GetView()));

            var menuTest = new ToolStripMenuItem("檢測數據");
            menuTest.DropDownItems.Add(CreateItem("G11 檢測數據看版", () => new App_TestDashboard().GetView()));
            menuTest.DropDownItems.Add(CreateItem("G12 量測項目一覽表", () => new App_TestMeasurementSummary().GetView())); 
            
            menuTest.DropDownItems.Add(CreateItem("G13 環測數據一覽表", () => new App_EnvTestSummary().GetView())); 
            
            menuTest.DropDownItems.Add(CreateItem("G19 檢測報告分析評估表", () => new App_TestReportEvaluation().GetView()));
            
            menuTest.DropDownItems.Add(new ToolStripSeparator());
            menuTest.DropDownItems.Add(CreateItem("G21 環境監測", () => new App_CoreTable("TestData", "EnvMonitor", "環境監測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G21.1 相似暴露族群劃分表", () => new App_CoreTable("TestData", "SimilarExposureGroup", "相似暴露族群劃分表", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G22 廢水定申檢", () => new App_CoreTable("TestData", "WastewaterPeriodic", "廢水定申檢", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G23 飲用水檢測", () => new App_CoreTable("TestData", "DrinkingWater", "飲用水檢測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G24 工業區檢驗", () => new App_CoreTable("TestData", "IndustrialZoneTest", "工業區檢驗", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G25 土壤氣體檢測", () => new App_CoreTable("TestData", "SoilGasTest", "土壤氣體檢測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G26 廢水自主檢驗", () => new App_CoreTable("TestData", "WastewaterSelfTest", "廢水自主檢驗", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G27 循環水檢測(廠商)", () => new App_CoreTable("TestData", "CoolingWaterVendor", "循環水檢測(廠商)", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G28 循環水檢測(自評)", () => new App_CoreTable("TestData", "CoolingWaterSelf", "循環水檢測(自評)", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G29 TCLP", () => new App_CoreTable("TestData", "TCLP", "TCLP毒性特性溶出", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("G30 水表校正", () => new App_CoreTable("TestData", "WaterMeterCalibration", "水錶校正", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(new ToolStripSeparator());
            menuTest.DropDownItems.Add(CreateItem("G31 其它檢測數據", () => new App_CoreTable("TestData", "OtherTests", "其它檢測數據", new DefaultLogic()).GetView()));

            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("H11 教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(new ToolStripSeparator());
            menuEdu.DropDownItems.Add(CreateItem("H21 訓練時數", () => new App_CoreTable("教育訓練", "訓練時數", "教育訓練時數", new DefaultLogic()).GetView()));

            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("I11 法規看板", () => new App_LawDashboard().GetView()));
            menuLaw.DropDownItems.Add(new ToolStripSeparator());
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "I21 環保法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "I22 職安衛法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "I23 消防法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "I24 其它法規"));

            var menuESG = new ToolStripMenuItem("ESG");
            menuESG.DropDownItems.Add(CreateItem("J11 ESG看板", () => new App_ESGDashboard().GetView())); 
            menuESG.DropDownItems.Add(new ToolStripSeparator());
            menuESG.DropDownItems.Add(CreateItem("J21 ESG績效管理", () => new App_CoreTable("ESG", "ESG_Performance", "ESG績效管理", new DefaultLogic()).GetView()));
            menuESG.DropDownItems.Add(CreateItem("J22 職業安全", () => new App_CoreTable("ESG", "ESG_OccupationalSafety", "職業安全", new DefaultLogic()).GetView()));
            menuESG.DropDownItems.Add(CreateItem("J23 健康衛生", () => new App_CoreTable("ESG", "ESG_HealthHygiene", "健康衛生", new DefaultLogic()).GetView()));
            menuESG.DropDownItems.Add(CreateItem("J24 環境與氣侯", () => new App_CoreTable("ESG", "ESG_EnvironmentClimate", "環境與氣侯", new DefaultLogic()).GetView()));
            menuESG.DropDownItems.Add(CreateItem("J25 消防與韌性", () => new App_CoreTable("ESG", "ESG_FireResilience", "消防與韌性", new DefaultLogic()).GetView()));

            var menuISO = new ToolStripMenuItem("ISO14001");
            menuISO.DropDownItems.Add(CreateItem("K11 ISO看板", () => new App_ISODashboard().GetView()));
            menuISO.DropDownItems.Add(new ToolStripSeparator());
            menuISO.DropDownItems.Add(CreateItem("K21 目標管理", () => new App_CoreTable("ISO14001", "TargetManagement", "目標管理", new DefaultLogic()).GetView()));
            
            var menuISOComm = new ToolStripMenuItem("環境溝通");
            menuISOComm.DropDownItems.Add(CreateItem("K31 環境資訊接收管制表", () => new App_CoreTable("ISO14001", "EnvInfoReceive", "環境資訊接收管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("K32 內文聯絡書管制表", () => new App_CoreTable("ISO14001", "InternalComm", "內文聯絡書管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("K33 郵件收文管制表", () => new App_CoreTable("ISO14001", "MailReceive", "郵件收文管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("K34 來賓拜訪紀錄表", () => new App_CoreTable("ISO14001", "VisitorRecord", "來賓拜訪紀錄表", new DefaultLogic()).GetView()));
            menuISO.DropDownItems.Add(menuISOComm);

            var menuApp = new ToolStripMenuItem("應用");
            LoadDynamicAppLinks(menuApp);

            menuApp.DropDownItems.Add(new ToolStripSeparator());
            var memReleaseItem = new ToolStripMenuItem("記憶體釋放");
            memReleaseItem.Click += (s, e) => MemoryOptimizer.Execute();
            menuApp.DropDownItems.Add(memReleaseItem);

            _menu1 = new ToolStripMenuItem("選單1") { Visible = false };   
            _menu1.DropDownItems.Add(CreateItem("WorkItems", () => new App_CoreTable("Menu1DB", "WorkItems", "WorkItems", new DefaultLogic()).GetView()));
            _menu1.DropDownItems.Add(CreateItem("統計看板", () => new App_StatsDashboard("Menu1DB").GetView())); // 🟢 註冊新的統計看板

            _menu2 = new ToolStripMenuItem("選單2") { Visible = false };
            _menu2.DropDownItems.Add(CreateItem("WorkItems", () => new App_CoreTable("Menu2DB", "WorkItems", "WorkItems", new DefaultLogic()).GetView()));

            _menu3 = new ToolStripMenuItem("選單3") { Visible = false };
            _menu3.DropDownItems.Add(CreateItem("WorkItems", () => new App_CoreTable("Menu3DB", "WorkItems", "WorkItems", new DefaultLogic()).GetView()));

            _menu4 = new ToolStripMenuItem("選單4") { Visible = false };
            _menu4.DropDownItems.Add(CreateItem("WorkItems", () => new App_CoreTable("Menu4DB", "WorkItems", "WorkItems", new DefaultLogic()).GetView()));

            var menuSettings = new ToolStripMenuItem("設定");

            var dbConfigItem = new ToolStripMenuItem("Z11 資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    string prompt = "進入設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                    if (AuthManager.VerifyAdmin(prompt)) { LoadModule(new App_DbConfig().GetView()); } 
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            var restoreDbItem = new ToolStripMenuItem("Z12 資料庫還原");
            restoreDbItem.Click += (s, e) => ShowDatabaseRestoreDialog();
            menuSettings.DropDownItems.Add(restoreDbItem);

            var menuManagerItem = new ToolStripMenuItem("Z13 選單管理 (自訂擴充)");
            menuManagerItem.Click += (s, e) => {
                string prompt = "進入設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (AuthManager.VerifyAdmin(prompt)) {
                    new App_MenuManager().ShowDialog(this);
                }
            };
            menuSettings.DropDownItems.Add(menuManagerItem);
            
            var dropdownItem = new ToolStripMenuItem("Z14 下拉選單與連動設定");
            dropdownItem.Click += (s, e) => {
                try {
                    new App_DropdownManager().ShowDialog(this);
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dropdownItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 

            var permissionItem = new ToolStripMenuItem("Z15 權限設定");
            permissionItem.Click += (s, e) => {
                string prompt = "管理系統權限需要系統管理者權限\n請輸入【Lv3系統管理者】\n密碼進行授權：";
                if (AuthManager.VerifyLv3Only(prompt)) {
                    new App_PermissionManager(_mainMenu).ShowDialog(this); // 🟢 修正的 _mainMenu 呼叫
                }
            };
            menuSettings.DropDownItems.Add(permissionItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 

            var reminderSettingItem = new ToolStripMenuItem("Z16 系統提醒設定");
            reminderSettingItem.Click += (s, e) => {
                try {
                    new App_ReminderManager().ShowDialog(this);
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(reminderSettingItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator());

            var appLinkSettingItem = new ToolStripMenuItem("Z17 應用連結設定");
            appLinkSettingItem.Click += (s, e) => {
                try {
                    new App_LinkManager().ShowDialog(this); 
                    LoadDynamicAppLinks(menuApp); 
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(appLinkSettingItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 

            var cleanupItem = new ToolStripMenuItem("Z18 附件檔案空間清理");
            cleanupItem.Click += (s, e) => {
                try {
                    string prompt = "執行空間清理需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                    if (AuthManager.VerifyAdmin(prompt)) { LoadModule(new App_AttachmentCleanup().GetView()); } 
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入清理模組：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(cleanupItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 

            var flowChartItem = new ToolStripMenuItem("Z19 系統流程圖");
            flowChartItem.Click += (s, e) => {
                try {
                    ForceEndCurrentEdit(); 
                    LoadModule(new App_SystemFlowchart().GetView());
                } catch (Exception ex) {
                    MessageBox.Show($"載入流程圖失敗：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(flowChartItem);

            var instructionItem = CreateItem("Z20 操作說明", () => new App_Instruction().GetView());
            menuSettings.DropDownItems.Add(instructionItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 
            
            var unlockMenuItem = new ToolStripMenuItem("Z21 開啟個人隱藏選單");
            unlockMenuItem.Click += UnlockMenu_Click;
            menuSettings.DropDownItems.Add(unlockMenuItem);

            var pwdMgmtItem = new ToolStripMenuItem("Z22 變更個人選單密碼");
            pwdMgmtItem.Click += (s, e) => {
                new App_PasswordManager().ShowDialog(this);
            };
            menuSettings.DropDownItems.Add(pwdMgmtItem);

            AttachCustomMenus(menuReports, menuSafety, menuChemical, menuChemReg, menuNursing, menuAir, menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuESG, menuISO, _menu1, _menu2, _menu3, _menu4);

            string currentUser = Environment.UserName.Trim();

            if (string.Equals(currentUser, "黃忠揚", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(currentUser, "TJ700657", StringComparison.OrdinalIgnoreCase))
            {
                _menu1.Visible = true;
            }
            else if (string.Equals(currentUser, "TJ700228", StringComparison.OrdinalIgnoreCase))
            {
                _menu2.Visible = true;
            }
            else if (string.Equals(currentUser, "TJ700533", StringComparison.OrdinalIgnoreCase))
            {
                _menu3.Visible = true;
            }
            else if (string.Equals(currentUser, "TJ204159", StringComparison.OrdinalIgnoreCase))
            {
                _menu4.Visible = true;
            }

            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuChemical, menuNursing, menuAirWaterWaste, 
                menuFire, menuTest, menuEdu, menuLaw, menuESG, menuISO, 
                menuApp, _menu1, _menu2, _menu3, _menu4, menuSettings 
            });

            ApplyViewPermissions();
        }

        private void ApplyViewPermissions()
        {
            try
            {
                string currentUser = Environment.UserName.Trim();
                HashSet<string> hiddenMenus = new HashSet<string>();

                string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                if (!File.Exists(dbPath)) return;

                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmdChk = new System.Data.SQLite.SQLiteCommand("CREATE TABLE IF NOT EXISTS [HiddenUserMenus] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [UserName] TEXT, [MenuText] TEXT, UNIQUE(UserName, MenuText));", conn)) {
                        cmdChk.ExecuteNonQuery();
                    }

                    using (var cmd = new System.Data.SQLite.SQLiteCommand("SELECT MenuText FROM HiddenUserMenus WHERE UserName=@U", conn))
                    {
                        cmd.Parameters.AddWithValue("@U", currentUser);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                hiddenMenus.Add(reader["MenuText"].ToString());
                            }
                        }
                    }
                }

                if (hiddenMenus.Count == 0) return; 

                HideMenuItems(_mainMenu.Items, hiddenMenus);
            }
            catch { }
        }

        private void HideMenuItems(ToolStripItemCollection items, HashSet<string> hiddenMenus)
        {
            foreach (ToolStripItem item in items)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    if (hiddenMenus.Contains(menuItem.Text))
                    {
                        menuItem.Visible = false;
                    }
                    else
                    {
                        HideMenuItems(menuItem.DropDownItems, hiddenMenus);
                    }
                }
            }
        }

        private class BackupItem 
        {
            public string Display { get; set; }
            public string Path { get; set; }
        }

        private void ShowDatabaseRestoreDialog()
        {
            if (!AuthManager.VerifyLv3Only("執行資料庫還原是極度危險操作！\n這將會覆蓋當前的資料。\n請輸入【Lv3系統管理者】密碼：")) 
                return;

            BackupManager.LoadConfig();
            string backupDir = BackupManager.BackupPath;

            if (!Directory.Exists(backupDir))
            {
                MessageBox.Show("找不到備份資料夾，無法進行還原！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var directories = new DirectoryInfo(backupDir).GetDirectories()
                                .OrderByDescending(d => d.CreationTime)
                                .ToList();

            if (directories.Count == 0)
            {
                MessageBox.Show("目前沒有任何可用的備份還原點！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (Form f = new Form())
            {
                f.Text = "⏳ 資料庫還原";
                f.Size = new Size(580, 520);
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.BackColor = Color.White;

                Label lblWarn = new Label { Text = "⚠️ 警告：還原後，目標資料將會被指定的備份時間點覆蓋！", ForeColor = Color.Crimson, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Location = new Point(25, 20), AutoSize = true };
                
                Label lblStep1 = new Label { Text = "1. 選擇要還原的備份時間點：", Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(25, 60), AutoSize = true };
                ComboBox cboTime = new ComboBox { Location = new Point(25, 85), Width = 500, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Consolas", 12F) };
                
                foreach (var dir in directories)
                {
                    string formattedName = dir.Name;
                    if (formattedName.Length == 13) 
                    {
                        formattedName = $"{formattedName.Substring(0,4)}年{formattedName.Substring(4,2)}月{formattedName.Substring(6,2)}日 {formattedName.Substring(9,2)}:{formattedName.Substring(11,2)}";
                    }
                    cboTime.Items.Add(new BackupItem { Display = formattedName, Path = dir.FullName });
                }
                cboTime.DisplayMember = "Display";
                cboTime.ValueMember = "Path";
                cboTime.SelectedIndex = 0;

                Label lblStep2 = new Label { Text = "2. 選擇還原範圍 (可指定特定資料庫與資料表)：", Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(25, 140), AutoSize = true };
                
                RadioButton rbAll = new RadioButton { Text = "🔥 災難還原 (覆蓋還原「全部」系統資料庫)", Location = new Point(40, 175), AutoSize = true, Checked = true, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.Crimson };
                RadioButton rbSpecific = new RadioButton { Text = "🎯 選擇性還原 (僅還原特定的資料庫或資料表)", Location = new Point(40, 210), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };

                Panel pnlSpecific = new Panel { Location = new Point(60, 245), Size = new Size(465, 110), BorderStyle = BorderStyle.FixedSingle, Enabled = false };
                
                Label lblDb = new Label { Text = "資料庫：", Location = new Point(15, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cboDb = new ComboBox { Location = new Point(115, 18), Width = 325, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };

                Label lblTb = new Label { Text = "資料表：", Location = new Point(15, 65), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                ComboBox cboTb = new ComboBox { Location = new Point(115, 63), Width = 325, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F) };
                cboTb.Items.Add(new App_DbConfig.ItemMap { EnName = "*", ChName = "-- 還原整個資料庫 --" });

                pnlSpecific.Controls.Add(lblDb);
                pnlSpecific.Controls.Add(cboDb);
                pnlSpecific.Controls.Add(lblTb);
                pnlSpecific.Controls.Add(cboTb);

                rbAll.CheckedChanged += (s, e) => pnlSpecific.Enabled = rbSpecific.Checked;
                
                var dbMap = App_DbConfig.GetDbMapCache();
                foreach (var kvp in dbMap) {
                    cboDb.Items.Add(new App_DbConfig.ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                }
                if (cboDb.Items.Count > 0) cboDb.SelectedIndex = 0;

                cboDb.SelectedIndexChanged += (s, e) => {
                    cboTb.Items.Clear();
                    cboTb.Items.Add(new App_DbConfig.ItemMap { EnName = "*", ChName = "-- 還原整個資料庫 --" });
                    
                    if (cboDb.SelectedItem is App_DbConfig.ItemMap map && !string.IsNullOrEmpty(map.EnName) && dbMap.ContainsKey(map.EnName)) {
                        foreach (var tbl in dbMap[map.EnName].Tables) {
                            cboTb.Items.Add(new App_DbConfig.ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                        }
                    }
                    cboTb.SelectedIndex = 0;
                };

                Button btnRestore = new Button { Text = "⚡ 立即執行還原並重啟系統", Location = new Point(170, 400), Size = new Size(240, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                
                btnRestore.Click += (s, e) => {
                    if (cboTime.SelectedItem == null) return;
                    BackupItem selectedTime = (BackupItem)cboTime.SelectedItem;
                    string sourceDir = selectedTime.Path;

                    string confirmMsg = "";
                    bool isFullRestore = rbAll.Checked;
                    string targetDb = "";
                    string targetTb = "";

                    if (isFullRestore) {
                        confirmMsg = $"您確定要將系統時光倒流至：\n\n【{selectedTime.Display}】嗎？\n\n(系統將自動關閉並覆蓋實體檔案)";
                    } else {
                        targetDb = ((App_DbConfig.ItemMap)cboDb.SelectedItem).EnName;
                        targetTb = ((App_DbConfig.ItemMap)cboTb.SelectedItem).EnName;
                        
                        string dbChName = ((App_DbConfig.ItemMap)cboDb.SelectedItem).ChName;
                        string tbChName = ((App_DbConfig.ItemMap)cboTb.SelectedItem).ChName;

                        if (targetTb == "*") {
                            confirmMsg = $"您即將還原單一資料庫：\n\n【{dbChName}】 ({targetDb})\n時間點：【{selectedTime.Display}】\n\n確定執行嗎？";
                        } else {
                            confirmMsg = $"您即將還原單一資料表：\n\n庫：{dbChName} ({targetDb})\n表：{tbChName} ({targetTb})\n時間點：【{selectedTime.Display}】\n\n(注意：系統將只抽取該表內容覆蓋當前資料庫)\n確定執行嗎？";
                        }
                    }

                    if (MessageBox.Show(confirmMsg, "最終確認", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
                    {
                        try
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            string destDir = DataManager.BasePath;

                            if (isFullRestore) 
                            {
                                string[] backupFiles = Directory.GetFiles(sourceDir, "*.sqlite");
                                foreach (var file in backupFiles)
                                {
                                    string destFile = Path.Combine(destDir, Path.GetFileName(file));
                                    File.Copy(file, destFile, true);
                                }
                                string sysConfigBackup = Path.Combine(sourceDir, "SystemConfig.sqlite");
                                if (File.Exists(sysConfigBackup)) {
                                    File.Copy(sysConfigBackup, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite"), true);
                                }
                            } 
                            else 
                            {
                                string sourceDbFile = Path.Combine(sourceDir, targetDb + ".sqlite");
                                string destDbFile = Path.Combine(destDir, targetDb + ".sqlite");

                                if (!File.Exists(sourceDbFile)) {
                                    MessageBox.Show("選擇的備份時間點中，找不到該資料庫檔案！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }

                                if (targetTb == "*") {
                                    File.Copy(sourceDbFile, destDbFile, true);
                                } else {
                                    string attachSql = $"ATTACH DATABASE '{sourceDbFile}' AS BackupDB;";
                                    string clearSql = $"DELETE FROM main.[{targetTb}];";
                                    string copySql = $"INSERT INTO main.[{targetTb}] SELECT * FROM BackupDB.[{targetTb}];";
                                    
                                    using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={destDbFile};Version=3;")) {
                                        conn.Open();
                                        using (var cmd = new System.Data.SQLite.SQLiteCommand(conn)) {
                                            cmd.CommandText = attachSql; cmd.ExecuteNonQuery();
                                            cmd.CommandText = clearSql; cmd.ExecuteNonQuery();
                                            cmd.CommandText = copySql; cmd.ExecuteNonQuery();
                                            cmd.CommandText = "DETACH DATABASE BackupDB;"; cmd.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }

                            MessageBox.Show("還原成功！系統將自動關閉，請重新啟動軟體。", "還原作業完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Environment.Exit(0);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"還原失敗：{ex.Message}\n可能是有其他同仁正在使用系統，導致檔案鎖定。請確保所有人都關閉軟體後再試一次。", "嚴重錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                };

                f.Controls.Add(lblWarn);
                f.Controls.Add(lblStep1);
                f.Controls.Add(cboTime);
                f.Controls.Add(lblStep2);
                f.Controls.Add(rbAll);
                f.Controls.Add(rbSpecific);
                f.Controls.Add(pnlSpecific);
                f.Controls.Add(btnRestore);
                f.ShowDialog();
            }
        }

        private void LoadDynamicAppLinks(ToolStripMenuItem menuApp)
        {
            menuApp.DropDownItems.Clear();
            
            var callExeItem = new ToolStripMenuItem("tgeOffice導入巡檢");
            callExeItem.Click += (s, e) => {
                string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"app\tgeoffice_dw\FormCrawlerApp.exe");
                try {
                    if (File.Exists(exePath)) Process.Start(new ProcessStartInfo { FileName = exePath, UseShellExecute = true });
                    else MessageBox.Show($"找不到外部程式：\n{exePath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex) { MessageBox.Show($"執行外部程式失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            menuApp.DropDownItems.Add(callExeItem);

            try {
                DataTable dt = DataManager.GetTableData("SystemConfig", "AppLinks", "", "", "");
                if (dt != null && dt.Rows.Count > 0)
                {
                    menuApp.DropDownItems.Add(new ToolStripSeparator());
                    foreach (DataRow row in dt.Rows)
                    {
                        string menuName = row["選單名稱"]?.ToString() ?? "未命名";
                        string exePath = row["執行路徑"]?.ToString() ?? "";

                        var customItem = new ToolStripMenuItem(menuName);
                        customItem.Click += (s, e) => {
                            try {
                                if (File.Exists(exePath)) Process.Start(new ProcessStartInfo { FileName = exePath, UseShellExecute = true });
                                else MessageBox.Show($"找不到外部程式：\n{exePath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            catch (Exception ex) { MessageBox.Show($"執行外部程式失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        };
                        menuApp.DropDownItems.Add(customItem);
                    }
                }
            } catch { }
        }

        private void AttachCustomMenus(params ToolStripMenuItem[] mainNodes)
        {
            try {
                DataTable dt = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
                if (dt == null) return;

                foreach (DataRow row in dt.Rows)
                {
                    string category = row["分類"].ToString();
                    string dbName = row["資料庫名"].ToString();
                    string tableName = row["資料表名"].ToString();

                    foreach (var node in mainNodes)
                    {
                        if (node.Text == category)
                        {
                            bool hasSeparator = false;
                            foreach (ToolStripItem item in node.DropDownItems) {
                                if (item is ToolStripSeparator && item.Tag != null && item.Tag.ToString() == "CustomDivider") {
                                    hasSeparator = true; break;
                                }
                            }
                            if (!hasSeparator) {
                                node.DropDownItems.Add(new ToolStripSeparator { Tag = "CustomDivider" });
                            }

                            node.DropDownItems.Add(CreateItem(tableName, () => new App_CoreTable(dbName, tableName, tableName, new DefaultLogic()).GetView()));
                            break;
                        }
                    }
                }
            } catch { }
        }

        private void UnlockMenu_Click(object sender, EventArgs e)
        {
            using (Form p = new Form())
            {
                p.Width = 400; 
                p.Height = 220;
                p.Text = "解鎖個人隱藏選單";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;

                Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入個人選單密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 250, Left = 30, Top = 70, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認解鎖", DialogResult = DialogResult.OK, Left = 160, Top = 120, Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor=Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog(this) == DialogResult.OK)
                {
                    string input = txt.Text.Trim();
                    string menuToUnlock = App_PasswordManager.CheckUnlockMenu(input);

                    if (menuToUnlock == "選單1") { _menu1.Visible = true; MessageBox.Show("「選單1」已解鎖！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    else if (menuToUnlock == "選單2") { _menu2.Visible = true; MessageBox.Show("「選單2」已解鎖！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    else if (menuToUnlock == "選單3") { _menu3.Visible = true; MessageBox.Show("「選單3」已解鎖！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    else if (menuToUnlock == "選單4") { _menu4.Visible = true; MessageBox.Show("「選單4」已解鎖！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    else
                    {
                        MessageBox.Show("密碼錯誤或無對應之個人選單！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += async (s, e) => {
                if (_contentPanel.Controls.Count > 0 && _contentPanel.Controls[0] is Label && _contentPanel.Controls[0].Text.Contains("載入中")) return;
                
                Application.UseWaitCursor = true;
                
                try {
                    ForceEndCurrentEdit();

                    _contentPanel.SuspendLayout();
                    _contentPanel.Controls.Clear();
                    Label lblLoading = new Label {
                        Text = $"⏳ 正在為您準備【{text}】的資料與畫面，請稍候...",
                        Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold),
                        ForeColor = Color.DimGray,
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter
                    };
                    _contentPanel.Controls.Add(lblLoading);
                    _contentPanel.ResumeLayout(true);

                    _contentPanel.Update(); 
                    Application.DoEvents(); 
                    
                    await Task.Delay(30);

                    Control view = getViewFunc(); 
                    if (view != null) {
                        LoadModule(view);
                    }
                } 
                catch (Exception ex) { 
                    MessageBox.Show($"載入模組 {text} 失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
                finally { 
                    Application.UseWaitCursor = false;
                }
            };
            return item;
        }

        private ToolStripMenuItem CreateLawItem(string dbName, string tableName)
        {
            var item = new ToolStripMenuItem(tableName);
            item.Click += async (s, e) => {
                if (_contentPanel.Controls.Count > 0 && _contentPanel.Controls[0] is Label && _contentPanel.Controls[0].Text.Contains("載入中")) return;

                Application.UseWaitCursor = true;

                try {
                    ForceEndCurrentEdit();
                    
                    _contentPanel.SuspendLayout();
                    _contentPanel.Controls.Clear();
                    Label lblLoading = new Label {
                        Text = $"⏳ 正在為您準備【{tableName}】的資料與畫面，請稍候...",
                        Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold),
                        ForeColor = Color.DimGray,
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter
                    };
                    _contentPanel.Controls.Add(lblLoading);
                    _contentPanel.ResumeLayout(true);

                    _contentPanel.Update();
                    Application.DoEvents();
                    await Task.Delay(30);

                    Control view = new App_CoreTable(dbName, tableName, tableName, new LawLogic()).GetView();
                    if (view != null) {
                        LoadModule(view);
                    }
                } 
                catch (Exception ex) { 
                    MessageBox.Show($"載入模組失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
                finally { 
                    Application.UseWaitCursor = false;
                }
            };
            return item;
        }

        public void LoadModule(Control moduleControl)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => LoadModule(moduleControl))); return; }
            if (moduleControl == null) return;

            try {
                _contentPanel.SuspendLayout();
                
                while (_contentPanel.Controls.Count > 0)
                {
                    Control ctrl = _contentPanel.Controls[0];
                    _contentPanel.Controls.Remove(ctrl);
                    ctrl.Dispose();
                }
                
                moduleControl.Dock = DockStyle.Fill;
                _contentPanel.Controls.Add(moduleControl);
                
                _contentPanel.ResumeLayout(true);
                
            } catch (Exception ex) {
                MessageBox.Show($"畫面切換時發生錯誤：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadWelcomeScreen()
        {
            ForceEndCurrentEdit();
            try { LoadModule(new App_HomeDashboard().GetView()); } 
            catch { }
        }
    }
}
