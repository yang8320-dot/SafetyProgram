using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices; // 必須引用以使用 Windows API

namespace Safety_System
{
    /// <summary>
    /// 主視窗：負責整體佈局與全域熱鍵喚醒
    /// </summary>
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        // --- 🟢 修正後的 Windows API 宣告 (加上了 static extern) ---
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9000;      // 自定義熱鍵 ID
        private const uint MOD_CONTROL = 0x0002; // Control 鍵代碼
        private const uint VK_3 = 0x33;          // 數字鍵 3 的代碼
        private const int WM_HOTKEY = 0x0312;    // 熱鍵訊息代碼

        public MainForm()
        {
            InitializeComponent();
            RegisterGlobalHotkey();
        }

        private void InitializeComponent()
        {
            // 視窗基本設定為 1440 x 810
            this.Text = "工安系統看板 (v4.7.2 - SQLite 版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            // 設定系統預設字體以確保清晰度
            this.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));

            // 初始化選單
            _mainMenu = new MenuStrip();
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

        // 註冊全域熱鍵 Ctrl + 3
        private void RegisterGlobalHotkey()
        {
            try
            {
                // 參數：視窗句柄, ID, 組合鍵, 主要鍵
                bool success = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_3);
                if (!success)
                {
                    // 若失敗通常是熱鍵被其他程式(如通訊軟體)佔用了
                }
            }
            catch { /* 防止環境不支持 API 時崩潰 */ }
        }

        // 監聽 Windows 訊息，當熱鍵觸發時呼叫
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                WakeUpWindow();
            }
            base.WndProc(ref m);
        }

        // 喚醒視窗邏輯
        private void WakeUpWindow()
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            this.Show();
            this.Activate();
            this.BringToFront();
        }

        private void BuildMenu()
        {
            // 1. 首頁
            var menuHome = new ToolStripMenuItem("首頁");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            // 2. 報表
            var menuReport = new ToolStripMenuItem("報表");
            var itemMonthly = new ToolStripMenuItem("月報表");
            itemMonthly.Click += (s, e) => LoadModule(new App_MonthlyReport().GetView());
            var itemYearly = new ToolStripMenuItem("年報表");
            itemYearly.Click += (s, e) => LoadModule(new App_YearlyReport().GetView());
            menuReport.DropDownItems.AddRange(new ToolStripItem[] { itemMonthly, itemYearly });

            // 3. 工安
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("虛驚", "App_NearMiss.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("工傷", "App_WorkInjury.cs"));
            menuSafety.DropDownItems.Add(CreatePlaceholderItem("交傷", "App_TrafficInjury.cs"));

            // 4. 環保、消防 (預留項目)
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(CreatePlaceholderItem("空污申報", "App_AirPollution.cs"));
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreatePlaceholderItem("消防設備", "App_FireEquip.cs"));

            // 5. 設定
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            itemDbConfig.Click += (s, e) => { new App_DbConfig().Show(); };
            menuSettings.DropDownItems.Add(itemDbConfig);

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), menuSettings
            });
        }

        private ToolStripMenuItem CreatePlaceholderItem(string text, string fileName)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => {
                _contentPanel.Controls.Clear();
                Label lbl = new Label {
                    Text = string.Format("本頁面對應 {0}, 資料尚在建立中", fileName),
                    Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Italic),
                    ForeColor = Color.DarkGray, AutoSize = true, Location = new Point(50, 50)
                };
                _contentPanel.Controls.Add(lbl);
            };
            return item;
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
            _contentPanel.Controls.Clear();
            Label lbl = new Label {
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n解析度：1440 x 810\n全域喚醒：Ctrl + 3",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }

        // 程式關閉時，釋放熱鍵資源
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { }
            base.OnFormClosing(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
