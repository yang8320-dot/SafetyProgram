/// FILE: Safety_System/MainForm.cs ///
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Safety_System
{
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        // 隱藏的個人選單
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
            this.Text = "工安系統看板 (v9.0 - 核心架構重構與極速效能版)";
            
            this.WindowState = FormWindowState.Maximized;
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);
            
            DataManager.LoadConfig();
            BackupManager.RunAutoBackup();
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
            
            LoadWelcomeScreen();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
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
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuReports = new ToolStripMenuItem("日常作業");
            menuReports.DropDownItems.Add(CreateItem("請購資料", () => new App_CoreTable("Purchase", "PurchaseData", "請購資料", new DefaultLogic()).GetView())); 
            menuReports.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReports.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(new ToolStripSeparator());
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄", () => new App_CoreTable("Safety", "SafetyInspection", "巡檢記錄管理", new DefaultLogic()).GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件", () => new App_CoreTable("Safety", "WorkInjury", "工傷事件管理", new DefaultLogic()).GetView()));
            menuSafety.DropDownItems.Add(CreateItem("輕傷事件", () => new App_CoreTable("Safety", "MinorInjury", "輕傷事件管理", new DefaultLogic()).GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交通意外", () => new App_CoreTable("Safety", "TrafficInjury", "交通意外紀錄", new DefaultLogic()).GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件", () => new App_CoreTable("Safety", "NearMiss", "虛驚事件管理", new DefaultLogic()).GetView()));
            menuSafety.DropDownItems.Add(CreateItem("安全觀察", () => new App_CoreTable("Safety", "SafetyObservation", "安全觀察紀錄", new DefaultLogic()).GetView()));

            var menuChemical = new ToolStripMenuItem("化學品");
            menuChemical.DropDownItems.Add(CreateItem("化學品看板", () => new App_ChemDashboard().GetView()));
            menuChemical.DropDownItems.Add(CreateItem("化學品快查", () => new App_ChemQuickSearch().GetView()));
            menuChemical.DropDownItems.Add(new ToolStripSeparator());

            var menuChemReg = new ToolStripMenuItem("化學品要求及規範");
            menuChemReg.DropDownItems.Add(CreateItem("1. 環測項目", () => new App_CoreTable("Chemical", "EnvTesting", "環測項目", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("2. 勞工暴露容許濃度", () => new App_CoreTable("Chemical", "ExposureLimits", "勞工暴露容許濃度", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("3. 毒性物質", () => new App_CoreTable("Chemical", "ToxicSubstances", "毒性物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("4. 關注性化學物質", () => new App_CoreTable("Chemical", "ConcernedChem", "關注性化學物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("5. 優先管理化學品", () => new App_CoreTable("Chemical", "PriorityMgmtChem", "優先管理化學品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("6. 管制化學品", () => new App_CoreTable("Chemical", "ControlledChem", "管制化學品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("7. 特定化學物質", () => new App_CoreTable("Chemical", "SpecificChem", "特定化學物質", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("8. 有機溶劑", () => new App_CoreTable("Chemical", "OrganicSolvents", "有機溶劑", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("9. 勞工健康保護", () => new App_CoreTable("Chemical", "WorkerHealthProtect", "勞工健康保護", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("10. 公共危險物品", () => new App_CoreTable("Chemical", "PublicHazardous", "公共危險物品", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("11. 空污緊急應變", () => new App_CoreTable("Chemical", "AirPollutionEmerg", "空污緊急應變", new DefaultLogic()).GetView()));
            menuChemReg.DropDownItems.Add(CreateItem("12. 工廠危險物品申報", () => new App_CoreTable("Chemical", "FactoryHazardous", "工廠危險物品申報", new DefaultLogic()).GetView()));
            menuChemical.DropDownItems.Add(menuChemReg);
            
            menuChemical.DropDownItems.Add(CreateItem("SDS清冊", () => new App_CoreTable("Chemical", "SDS_Inventory", "SDS清冊", new DefaultLogic()).GetView()));

            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(new ToolStripSeparator());
            menuNursing.DropDownItems.Add(CreateItem("健康促進活動", () => new App_CoreTable("Nursing", "HealthPromotion", "健康促進活動", new DefaultLogic()).GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報紀錄", () => new App_CoreTable("Nursing", "WorkInjuryReport", "職災申報紀錄", new DefaultLogic()).GetView()));

            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(new ToolStripSeparator());
            menuAir.DropDownItems.Add(CreateItem("空污申報紀錄", () => new App_CoreTable("Air", "AirPollution", "空污申報紀錄", new DefaultLogic()).GetView()));

            var menuWater = new ToolStripMenuItem("水污");
            menuWater.DropDownItems.Add(CreateItem("水資源管理看板", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(new ToolStripSeparator());
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理水量記錄", () => new App_CoreTable("Water", "WaterMeterReadings", "【日】廢水處理水量記錄", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理用藥記錄", () => new App_CoreTable("Water", "WaterChemicals", "【日】廢水處理用藥記錄", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】自來水使用量", () => new App_CoreTable("Water", "WaterUsageDaily", "【日】自來水使用量", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】納管排放數據", () => new App_CoreTable("Water", "DischargeData", "【月】納管排放數據", new WaterLogic()).GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】自來水用量統計", () => new App_CoreTable("Water", "WaterVolume", "【月】自來水用量統計", new WaterLogic()).GetView()));

            var menuWaste = new ToolStripMenuItem("產能及廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("產能及廢棄物看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(new ToolStripSeparator());
            menuWaste.DropDownItems.Add(CreateItem("【月】複層月表", () => new App_CoreTable("Waste", "Waste_IL", "【月】複層月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】膠合月表", () => new App_CoreTable("Waste", "Waste_LM", "【月】膠合月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】鍍板月表", () => new App_CoreTable("Waste", "Waste_CR", "【月】鍍板月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】強化月表", () => new App_CoreTable("Waste", "Waste_T", "【月】強化月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】切磨月表", () => new App_CoreTable("Waste", "Waste_GCTE", "【月】切磨月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】物料月表", () => new App_CoreTable("Waste", "Waste_ML", "【月】物料月表", new DefaultLogic()).GetView()));
            menuWaste.DropDownItems.Add(CreateItem("【月】水站月表", () => new App_CoreTable("Waste", "Waste_Water", "【月】水站月表", new DefaultLogic()).GetView()));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(new ToolStripSeparator());
            menuFire.DropDownItems.Add(CreateItem("火源責任人管理", () => new App_CoreTable("Fire", "FireResponsible", "火源責任人管理", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_CoreTable("Fire", "HazardStats", "公共危險物統計", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_CoreTable("Fire", "FireEquip", "消防設備巡檢", new DefaultLogic()).GetView()));
            menuFire.DropDownItems.Add(CreateItem("各單位消防自主檢查表", () => new App_CoreTable("Fire", "FireSelfInspection", "各單位消防自主檢查表", new DefaultLogic()).GetView()));

            var menuTest = new ToolStripMenuItem("檢測數據");
            menuTest.DropDownItems.Add(CreateItem("檢測數據看版", () => new App_TestDashboard().GetView()));
            menuTest.DropDownItems.Add(new ToolStripSeparator());
            menuTest.DropDownItems.Add(CreateItem("環境監測", () => new App_CoreTable("TestData", "EnvMonitor", "環境監測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水定申檢", () => new App_CoreTable("TestData", "WastewaterPeriodic", "廢水定申檢", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("飲用水檢測", () => new App_CoreTable("TestData", "DrinkingWater", "飲用水檢測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("工業區檢驗", () => new App_CoreTable("TestData", "IndustrialZoneTest", "工業區檢驗", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("土壤氣體檢測", () => new App_CoreTable("TestData", "SoilGasTest", "土壤氣體檢測", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水自主檢驗", () => new App_CoreTable("TestData", "WastewaterSelfTest", "廢水自主檢驗", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(廠商)", () => new App_CoreTable("TestData", "CoolingWaterVendor", "循環水檢測(廠商)", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(自評)", () => new App_CoreTable("TestData", "CoolingWaterSelf", "循環水檢測(自評)", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("TCLP", () => new App_CoreTable("TestData", "TCLP", "TCLP毒性特性溶出", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("水錶校正", () => new App_CoreTable("TestData", "WaterMeterCalibration", "水錶校正", new DefaultLogic()).GetView()));
            menuTest.DropDownItems.Add(CreateItem("其它檢測數據", () => new App_CoreTable("TestData", "OtherTests", "其它檢測數據", new DefaultLogic()).GetView()));

            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(new ToolStripSeparator());
            menuEdu.DropDownItems.Add(CreateItem("訓練時數", () => new App_CoreTable("教育訓練", "訓練時數", "教育訓練時數", new DefaultLogic()).GetView()));

            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("法規看板", () => new App_LawDashboard().GetView()));
            menuLaw.DropDownItems.Add(new ToolStripSeparator());
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "環保法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "職安衛法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "消防法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "其它法規"));

            var menuESG = new ToolStripMenuItem("ESG");
            menuESG.DropDownItems.Add(CreateItem("ESG看板", () => new App_ESGDashboard().GetView())); 
            menuESG.DropDownItems.Add(new ToolStripSeparator());
            menuESG.DropDownItems.Add(CreateItem("ESG績效管理", () => new App_CoreTable("ESG", "ESG_Performance", "ESG績效管理", new DefaultLogic()).GetView()));

            var menuISO = new ToolStripMenuItem("ISO14001");
            menuISO.DropDownItems.Add(CreateItem("ISO看板", () => new App_ISODashboard().GetView()));
            menuISO.DropDownItems.Add(new ToolStripSeparator());
            menuISO.DropDownItems.Add(CreateItem("目標管理", () => new App_CoreTable("ISO14001", "TargetManagement", "目標管理", new DefaultLogic()).GetView()));
            
            var menuISOComm = new ToolStripMenuItem("環境溝通");
            menuISOComm.DropDownItems.Add(CreateItem("環境資訊接收管制表", () => new App_CoreTable("ISO14001", "EnvInfoReceive", "環境資訊接收管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("內文聯絡書管制表", () => new App_CoreTable("ISO14001", "InternalComm", "內文聯絡書管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("郵件收文管制表", () => new App_CoreTable("ISO14001", "MailReceive", "郵件收文管制表", new DefaultLogic()).GetView()));
            menuISOComm.DropDownItems.Add(CreateItem("來賓拜訪紀錄表", () => new App_CoreTable("ISO14001", "VisitorRecord", "來賓拜訪紀錄表", new DefaultLogic()).GetView()));
            menuISO.DropDownItems.Add(menuISOComm);

            var menuApp = new ToolStripMenuItem("應用");
            var callExeItem = new ToolStripMenuItem("tgeOffice導入巡檢");
            callExeItem.Click += (s, e) => {
                string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"app\tgeoffice_dw\FormCrawlerApp.exe");
                try
                {
                    if (File.Exists(exePath))
                    {
                        Process.Start(new ProcessStartInfo { FileName = exePath, UseShellExecute = true });
                    }
                    else
                    {
                        MessageBox.Show($"找不到外部程式：\n{exePath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"執行外部程式失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuApp.DropDownItems.Add(callExeItem);

            _menu1 = new ToolStripMenuItem("選單1") { Visible = false };   
            _menu1.DropDownItems.Add(CreateItem("KPI", () => new App_CoreTable("Menu1DB", "KPI", "KPI", new DefaultLogic()).GetView()));
            _menu1.DropDownItems.Add(CreateItem("文化改善", () => new App_CoreTable("Menu1DB", "CultureImprove", "文化改善", new DefaultLogic()).GetView()));
            _menu1.DropDownItems.Add(CreateItem("PBC", () => new App_CoreTable("Menu1DB", "PBC", "PBC", new DefaultLogic()).GetView()));
            _menu1.DropDownItems.Add(CreateItem("帳密管理", () => new App_CoreTable("Menu1DB", "AccountManage", "帳密管理", new DefaultLogic()).GetView()));

            _menu2 = new ToolStripMenuItem("選單2") { Visible = false };
            _menu2.DropDownItems.Add(CreateItem("資料管理", () => new App_CoreTable("Menu2DB", "DataManage2", "資料管理", new DefaultLogic()).GetView()));

            _menu3 = new ToolStripMenuItem("選單3") { Visible = false };
            _menu3.DropDownItems.Add(CreateItem("資料管理3", () => new App_CoreTable("Menu3DB", "DataManage3", "資料管理3", new DefaultLogic()).GetView()));

            _menu4 = new ToolStripMenuItem("選單4") { Visible = false };
            _menu4.DropDownItems.Add(CreateItem("資料管理4", () => new App_CoreTable("Menu4DB", "DataManage4", "資料管理4", new DefaultLogic()).GetView()));

            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));

            var menuManagerItem = new ToolStripMenuItem("選單管理 (自訂擴充)");
            menuManagerItem.Click += (s, e) => {
                new App_MenuManager().ShowDialog(this);
            };
            menuSettings.DropDownItems.Add(menuManagerItem);
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    if (AuthManager.VerifyAdmin()) { LoadModule(new App_DbConfig().GetView()); } 
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            var cleanupItem = new ToolStripMenuItem("附件檔案空間清理");
            cleanupItem.Click += (s, e) => {
                try {
                    if (AuthManager.VerifyAdmin("執行空間清理需要管理者權限，請輸入密碼：")) { LoadModule(new App_AttachmentCleanup().GetView()); } 
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入清理模組：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(cleanupItem);

            menuSettings.DropDownItems.Add(new ToolStripSeparator()); 
            
            var unlockMenuItem = new ToolStripMenuItem("開啟個人選單");
            unlockMenuItem.Click += UnlockMenu_Click;
            menuSettings.DropDownItems.Add(unlockMenuItem);

            var pwdMgmtItem = new ToolStripMenuItem("密碼管理");
            pwdMgmtItem.Click += (s, e) => {
                new App_PasswordManager().ShowDialog(this);
            };
            menuSettings.DropDownItems.Add(pwdMgmtItem);

            AttachCustomMenus(menuReports, menuSafety, menuChemical, menuChemReg, menuNursing, menuAir, menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuESG, menuISO, _menu1, _menu2, _menu3, _menu4);

            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuChemical, menuNursing, menuAir, 
                menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuESG, menuISO, 
                menuApp, _menu1, _menu2, _menu3, _menu4, menuSettings 
            });
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
                p.Text = "解鎖個人選單";
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
                try { LoadModule(new App_CoreTable(dbName, tableName, tableName, new LawLogic()).GetView()); } 
                catch (Exception ex) { MessageBox.Show($"載入模組失敗：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            return item;
        }

        public void LoadModule(Control moduleControl)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => LoadModule(moduleControl))); return; }
            if (moduleControl == null) return;
            try {
                while (_contentPanel.Controls.Count > 0)
                {
                    Control ctrl = _contentPanel.Controls[0];
                    _contentPanel.Controls.Remove(ctrl);
                    ctrl.Dispose();
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();

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
