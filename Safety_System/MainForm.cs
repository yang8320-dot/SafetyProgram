using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 主視窗：負責整體佈局與模組畫面切換
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
            // 視窗基本設定為 1440 x 810
            this.Text = "工安系統看板 (v4.7.2)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            // 🟢 關鍵修正 1.2: 將系統預設字體改為較銳利的 UI 字體
            this.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));

            // 🟢 關鍵修正 2.1: 確保視窗啟用快捷鍵預覽
            this.KeyPreview = true;

            // 初始化選單
            _mainMenu = new MenuStrip();
            // 選單字體也要單獨設定以保持一致
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

            // 主顯示容器 (大框)
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30),
                AutoScroll = true
            };
            this.Controls.Add(_contentPanel);
            
            _mainMenu.BringToFront();
            LoadWelcomeScreen();
        }

        private void BuildMenu()
        {
            // 工安選單
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("虛驚"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("工傷"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("交傷"));

            // 環保選單
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(new ToolStripMenuItem("空污申報"));

            // 消防選單
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(new ToolStripMenuItem("消防設備"));

            // 設定選單
            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(new ToolStripMenuItem("使用者設定"));

            // 將所有選單加入列
            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                menuSettings
            });
        }

        /// <summary>
        /// 🟢 關鍵修正 2.2: 實際處理 Ctrl + 3 快捷鍵並載入畫面
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.D3) || keyData == (Keys.Control | Keys.NumPad3))
            {
                // 當按下了 Ctrl + 3 (主鍵盤或數字鍵盤)，直接載入「巡檢紀錄」畫面
                LoadModule(new App_SafetyInspection().GetView());
                return true; // 代表按鍵事件已處理，不要再往下傳遞
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

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
            // 使用 UI 字體設定歡迎文字
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n架構：.NET Framework 4.7.2\n(快捷鍵: Ctrl+3 快速開啟巡檢)",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
