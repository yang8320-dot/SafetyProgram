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

        // ==========================================
        // 🟢 全域快捷鍵攔截邏輯 (Ctrl + S 存檔)
        // ==========================================
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                // 在內容面板中尋找名稱為 "btnSave" 的按鈕
                Button activeSaveButton = FindControlByName(_contentPanel, "btnSave") as Button;

                if (activeSaveButton != null && activeSaveButton.Enabled)
                {
                    // 強制結束所有控制項的輸入狀態，確保 DataGridView 的最後一格資料被確認
                    this.Validate(); 
                    
                    // 模擬點擊儲存按鈕
                    activeSaveButton.PerformClick();
                    return true; // 攔截按鍵，防止發出系統警告音
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // 🟢 輔助方法：遞迴尋找控制項
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
        // ==========================================

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
            menuTest.DropDownItems.Add(CreateItem("環境監測", () => new App_EnvMonitor().GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水定申檢", () => new App_WastewaterPeriodic().GetView()));
            menuTest.DropDownItems.Add(CreateItem("飲用水檢測", () => new App_DrinkingWater().GetView()));
            menuTest.DropDownItems.Add(CreateItem("工業區檢驗", () => new App_IndustrialZoneTest().GetView()));
            menuTest.DropDownItems.Add(CreateItem("土壤氣體檢測", () => new App_SoilGasTest().GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水自主檢驗", () => new App_WastewaterSelfTest().GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(廠商)", () => new App_CoolingWaterVendor().GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(自評)", () => new App_CoolingWaterSelf().GetView()));
            menuTest.DropDownItems.Add(CreateItem("TCLP", () => new App_TCLP().GetView()));
            menuTest.DropDownItems.Add(CreateItem("水錶校正", () => new App_WaterMeterCalibration().GetView()));
            menuTest.DropDownItems.Add(CreateItem("其它檢測數據", () => new App_OtherTests().GetView()));

            // ================= 新增：教育訓練 =================
            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(CreateItem("訓練時數", () => new App_EduHours().GetView()));

            // ================= 新增：法規 (動態載入共用模組) =================
            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("法規看板", () => new App_LawDashboard().GetView()));

            // 1. 環保法規 (20項)
            var menuLawEnv = new ToolStripMenuItem("環保法規");
            string[] envLaws = { "組織與處務", "環境綜合計畫", "環境影響評估", "空氣污染防制", "噪音污染管制", "水污染防治", "海洋污染防制", "廢棄物清理", "應回收廢棄物", "資源回收再利用", "土壤污染整治", "毒化物管理", "飲用水管理", "環境用藥管理", "公害糾紛處理", "環境污染檢驗", "環保人員訓練", "環境教育", "溫室氣體管理", "室內空氣管理" };
            foreach (var law in envLaws) menuLawEnv.DropDownItems.Add(CreateLawItem("環保法規", law));

            // 2. 職安衛法規 (17項)
            var menuLawOsh = new ToolStripMenuItem("職安衛法規");
            string[] oshLaws = { "一般安全衛生法規", "一般環境管理相關法規", "高壓氣體相關法規", "健康管理相關法規", "教育訓練相關法規", "化學物質相關法規", "機械安全相關法規", "特殊作業相關法規", "特別行業適用法規", "職業災害勞工保護法相關法規", "安衛其他法規", "勞資關係", "勞動條件", "勞工福利", "勞工保險", "職業訓練", "就業服務" };
            foreach (var law in oshLaws) menuLawOsh.DropDownItems.Add(CreateLawItem("職安衛法規", law));

            // 3. 其他法規 (25項)
            var menuLawOther = new ToolStripMenuItem("其他法規");
            string[] otherLaws = { "原子能相關法規", "衛生福利相關法規", "交通安全相關法規", "消防相關法規", "建築相關法規", "下水道相關法規", "科學園區相關法規", "一般工業區相關法規", "利害相關者要求", "國際環保公約", "水利相關法規", "能源管理", "文化相關法規", "族群相關法規", "消費相關法規", "財政金融相關法規", "法務相關法規", "社福警政相關法規", "經濟相關法規", "觀光旅遊相關法規", "生態相關法規", "礦業相關法規", "農林漁牧相關法規", "教育體育相關法規", "通訊傳播相關法規" };
            foreach (var law in otherLaws) menuLawOther.DropDownItems.Add(CreateLawItem("其它法規", law));

            menuLaw.DropDownItems.Add(menuLawEnv);
            menuLaw.DropDownItems.Add(menuLawOsh);
            menuLaw.DropDownItems.Add(menuLawOther);

            // ================= 原有：設定 =================
            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    if (VerifyAdminPassword()) { LoadModule(new App_DbConfig().GetView()); } 
                    else { MessageBox.Show("密碼錯誤，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            // 🟢 將教育訓練(menuEdu)與法規(menuLaw)加入主選單
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

        // 🟢 新增：專門用來產生共用法規模組的 Helper
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

        private bool VerifyAdminPassword()
        {
            Form p = new Form { Width = 400, Height = 230, Text = "管理員驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label l = new Label { Text = "請輸入系統管理員密碼：", Left = 20, Top = 20, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Left = 20, Top = 60, Width = 340, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", Left = 260, Top = 120, Width = 100, Height = 35, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.Add(l); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }

        private void LoadWelcomeScreen()
        {
            try { LoadModule(new App_HomeDashboard().GetView()); } 
            catch { }
        }
    }
}
