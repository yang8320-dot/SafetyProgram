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
            this.Text = "工安系統看板 (v5.0.0 - 佈局與資料升級版)";
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
                // 🟢 需求 1：主選單與下方頁面間隔設為 30，左右下邊框緩衝為 15
                Padding = new Padding(15, 30, 15, 15),
                AutoScroll = true
            };

            this.Controls.Add(_contentPanel);
            this.Controls.Add(_mainMenu); 
            
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // (因篇幅省略各模組細節，僅保留核心架構與設定選單)
            var menuSafety = new ToolStripMenuItem("工安管理");
            menuSafety.DropDownItems.Add(CreateItem("工安管理看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件管理", () => new App_NearMiss().GetView()));

            // 🟢 需求 3.1：資料庫設定加入密碼保護
            var menuSettings = new ToolStripMenuItem("系統設定");
            menuSettings.DropDownItems.Add(CreateItem("月報統計", () => new App_MonthlyReport().GetView()));
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                if (VerifyAdminPassword()) {
                    LoadModule(new App_DbConfig().GetView()); // 整合為 GetView 模式
                } else {
                    MessageBox.Show("密碼錯誤，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            _mainMenu.Items.AddRange(new ToolStripItem[] { menuSafety, menuSettings });
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

        // 🟢 密碼驗證邏輯
        private bool VerifyAdminPassword()
        {
            Form p = new Form { Width = 400, Height = 200, Text = "管理員驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label l = new Label { Text = "請輸入系統管理員密碼：", Left = 20, Top = 20, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Left = 20, Top = 60, Width = 340, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", Left = 260, Top = 110, Width = 100, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 12F) };
            p.Controls.Add(l); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces"; // 密碼 tces
        }

        private void LoadWelcomeScreen()
        {
            _contentPanel.Controls.Clear();
            Label lbl = new Label { Text = "=== 工安系統看板 ===\n已套用：排版優化、日期格式化與防重寫機制", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(50, 50) };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
