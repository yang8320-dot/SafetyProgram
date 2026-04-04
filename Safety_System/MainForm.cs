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
            this.Text = "工安系統看板 (v4.8.9 - 修正編譯版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1280, 720);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            _mainMenu.Dock = DockStyle.Top;
            BuildMenu();

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(20),
                AutoScroll = true
            };

            this.Controls.Add(_contentPanel);
            this.Controls.Add(_mainMenu);
            
            this.Load += (s, e) => LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuWater = new ToolStripMenuItem("水");
            menuWater.DropDownItems.Add(CreateItem("水處理記錄表", () => new App_WaterTreatment().GetView()));

            var menuSettings = new ToolStripMenuItem("設定");
            
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            // 🟢 修正：使用 LoadModule 將 GetView() 回傳的控制項嵌入面板，不再使用 .Show()
            itemDbConfig.Click += (s, e) => LoadModule(new App_DbConfig().GetView());

            var itemInstruction = new ToolStripMenuItem("說明");
            // itemInstruction.Click += (s, e) => LoadModule(new App_Instruction().GetView());

            menuSettings.DropDownItems.AddRange(new ToolStripItem[] { itemDbConfig, itemInstruction });

            _mainMenu.Items.AddRange(new ToolStripItem[] { menuHome, menuWater, menuSettings });
        }

        private ToolStripMenuItem CreateItem(string text, Func<Control> getViewFunc)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                try { LoadModule(getViewFunc()); }
                catch (Exception ex) { MessageBox.Show($"無法載入模組：{ex.Message}"); }
            };
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
            Label lbl = new Label {
                Text = "=== 工安系統看板 ===\n編譯狀態：已修復\n資料引擎：SQLite (自動覆寫重複日期)",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, 
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
