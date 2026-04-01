using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 主視窗：定義系統解析度為 1440x810 並負責模組切換
    /// </summary>
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
            // 1. 視窗基本設定 (解析度調整)
            this.Text = "工安系統看板";
            this.Size = new Size(1440, 810); 
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            // 2. 初始化選單
            _mainMenu = new MenuStrip();
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 3. 初始化主內容容器
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30), // 加大內距，適合寬螢幕佈局
                AutoScroll = true          // 若內容過多自動顯示捲軸
            };
            this.Controls.Add(_contentPanel);
            
            // 確保選單在最上層
            _mainMenu.BringToFront();

            // 4. 載入初始畫面
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // 工安選單
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            menuSafety.DropDownItems.AddRange(new ToolStripItem[] {
                itemInspection,
                new ToolStripMenuItem("虛驚"),
                new ToolStripMenuItem("工傷"),
                new ToolStripMenuItem("交傷")
            });

            // 環保選單
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(new ToolStripMenuItem("空污申報"));

            // 消防選單
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(new ToolStripMenuItem("消防設備"));

            // 其他選單
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                new ToolStripMenuItem("設定")
            });
        }

        /// <summary>
        /// 切換中間內容面板的邏輯
        /// </summary>
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
            Label lblWelcome = new Label
            {
                Text = "=== 工安系統看板 ===\n\n系統環境已就緒，請由上方選單選擇功能項目。",
                Font = new Font("Microsoft JhengHei", 24, FontStyle.Bold), // 大解析度下使用較大字體
                AutoSize = true,
                Location = new Point(50, 50)
            };
            
            Panel pnlWelcome = new Panel { Dock = DockStyle.Fill };
            pnlWelcome.Controls.Add(lblWelcome);
            LoadModule(pnlWelcome);
        }
    }
}
