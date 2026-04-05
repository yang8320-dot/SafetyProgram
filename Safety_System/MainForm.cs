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
            menuWater.DropDownItems.Add(CreateItem("納管排放數據", () => new App_DischargeData().GetView()));
            menuWater.DropDownItems.Add(CreateItem("廢水處理水量記錄", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("廢水處理用藥記錄", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("自來水用量統計", () => new App_WaterVolume().GetView()));

            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("廢棄物統計表", () => new App_WasteMonthly().GetView()));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人管理", () => new App_FireResponsible().GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_HazardStats().GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_FireEquip().GetView()));

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

            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuNursing, menuAir, 
                menuWater, menuWaste, menuFire, menuSettings 
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
            // 🟢 修正：高度增加至 230，確保按鈕不被遮擋
            Form p = new Form { Width = 400, Height = 230, Text = "管理員驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label l = new Label { Text = "請輸入系統管理員密碼：", Left = 20, Top = 20, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Left = 20, Top = 60, Width = 340, Font = new Font("Microsoft JhengHei UI", 14F) };
            
            // 🟢 修正：按鈕 Top 調整至 120，留出更多空間
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
