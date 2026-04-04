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
            this.Text = "工安系統看板 (v5.0.1 - 選單完整還原版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip { Font = new Font("Microsoft JhengHei UI", 12F), Dock = DockStyle.Top };
            BuildMenu();

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                // 🟢 需求：主選單與下方頁面間隔設為 30 (Top)，左右下邊框緩衝為 15 (框內與文字間隔)
                Padding = new Padding(15, 30, 15, 15),
                AutoScroll = true
            };

            this.Controls.Add(_contentPanel);
            this.Controls.Add(_mainMenu); 
            
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // 1. 工安管理
            var menuSafety = new ToolStripMenuItem("工安管理");
            menuSafety.DropDownItems.Add(CreateItem("工安管理看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件管理", () => new App_NearMiss().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄管理", () => new App_SafetyInspection().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("安全觀察紀錄", () => new App_SafetyObservation().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交通意外紀錄", () => new App_TrafficInjury().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件管理", () => new App_WorkInjury().GetView()));

            // 2. 職場護理
            var menuNursing = new ToolStripMenuItem("職場護理");
            menuNursing.DropDownItems.Add(CreateItem("職場健康看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進活動", () => new App_HealthPromotion().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報紀錄", () => new App_WorkInjuryReport().GetView()));

            // 3. 環境保護 - 空污
            var menuAir = new ToolStripMenuItem("空污防治");
            menuAir.DropDownItems.Add(CreateItem("空氣汙染看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報紀錄", () => new App_AirPollution().GetView()));

            // 4. 環境保護 - 水質
            var menuWater = new ToolStripMenuItem("水資源管理");
            menuWater.DropDownItems.Add(CreateItem("水資源看板", () => new App_WaterDashboard().GetView()));
            menuWater.DropDownItems.Add(CreateItem("納管排放數據", () => new App_DischargeData().GetView()));
            menuWater.DropDownItems.Add(CreateItem("水處理記錄", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("水處理用藥記錄", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("用水量統計", () => new App_WaterVolume().GetView()));

            // 5. 環境保護 - 廢棄物
            var menuWaste = new ToolStripMenuItem("廢棄物管理");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物清運看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("廢棄物月報", () => new App_WasteMonthly().GetView()));

            // 6. 消防安全
            var menuFire = new ToolStripMenuItem("消防安全");
            menuFire.DropDownItems.Add(CreateItem("消防安全看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人", () => new App_FireResponsible().GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_HazardStats().GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_FireEquip().GetView()));

            // 7. 系統設定 (含密碼保護的資料庫設定)
            var menuSettings = new ToolStripMenuItem("系統設定");
            menuSettings.DropDownItems.Add(CreateItem("月度統計報表", () => new App_MonthlyReport().GetView()));
            menuSettings.DropDownItems.Add(CreateItem("年度績效報表", () => new App_YearlyReport().GetView()));
            menuSettings.DropDownItems.Add(CreateItem("系統操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫與防重寫設定");
            dbConfigItem.Click += (s, e) => {
                if (VerifyAdminPassword()) {
                    LoadModule(new App_DbConfig().GetView());
                } else {
                    MessageBox.Show("密碼錯誤，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            // 將所有主選單加入 MenuStrip
            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuSafety, menuNursing, menuAir, menuWater, menuWaste, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings 
            });
        }

        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                try { LoadModule(getViewFunc()); } 
                catch (Exception ex) { MessageBox.Show($"無法載入 {text}：\n{ex.Message}"); }
            };
            return item;
        }

        public void LoadModule(Control moduleControl)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => LoadModule(moduleControl))); return; }
            _contentPanel.Controls.Clear();
            moduleControl.Dock = DockStyle.Fill;
            _contentPanel.Controls.Add(moduleControl);
        }

        // 🟢 管理員密碼驗證邏輯 (密碼: tces)
        private bool VerifyAdminPassword()
        {
            Form p = new Form { Width = 400, Height = 200, Text = "管理員驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label l = new Label { Text = "請輸入系統管理員密碼：", Left = 20, Top = 20, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Left = 20, Top = 60, Width = 340, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", Left = 260, Top = 110, Width = 100, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.Add(l); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }

        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label { 
                Text = "=== 工安系統看板 ===\n已套用：選單完整還原、排版優化與防重寫機制", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                AutoSize = true, 
                Location = new Point(50, 50) 
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
