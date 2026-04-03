using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Safety_System
{
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9000;
        private const uint MOD_CONTROL = 0x0002;
        private const uint VK_3 = 0x33;
        private const int WM_HOTKEY = 0x0312;

        public MainForm()
        {
            InitializeComponent();
            RegisterGlobalHotkey();
        }

        private void InitializeComponent()
        {
            this.Text = "工安系統看板 (v4.7.2 - SQLite 版)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _mainMenu = new MenuStrip();
            _mainMenu.Font = new Font("Microsoft JhengHei UI", 12F);
            BuildMenu();
            this.Controls.Add(_mainMenu);

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

            // 4. 環保
            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(CreatePlaceholderItem("空污申報", "App_AirPollution.cs"));

            // 5. 水資源
            var menuWater = new ToolStripMenuItem("水資源");
            var itemWaterTreat = new ToolStripMenuItem("水處理記錄表");
            itemWaterTreat.Click += (s, e) => LoadModule(new App_WaterTreatment().GetView());
            
            var itemWaterChem = new ToolStripMenuItem("用藥記錄表");
            itemWaterChem.Click += (s, e) => LoadModule(new App_WaterChemicals().GetView());
            
            var itemWaterVol = new ToolStripMenuItem("自水水量");
            itemWaterVol.Click += (s, e) => LoadModule(new App_WaterVolume().GetView());

            // 🟢 新增：納管排放數據
            var itemDischarge = new ToolStripMenuItem("納管排放數據");
            itemDischarge.Click += (s, e) => LoadModule(new App_DischargeData().GetView());
            
            menuWater.DropDownItems.AddRange(new ToolStripItem[] { itemWaterTreat, itemWaterChem, itemWaterVol, itemDischarge });

            // 6. 消防
            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreatePlaceholderItem("消防設備", "App_FireEquip.cs"));

            // 7. 設定
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("資料庫設定");
            itemDbConfig.Click += (s, e) => { new App_DbConfig().Show(); };
            menuSettings.DropDownItems.Add(itemDbConfig);

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuHome, menuReport, menuSafety, menuEnv, menuWater, menuFire, 
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
                Text = "=== 工安系統看板 ===\n資料引擎：SQLite\n全域喚醒：Ctrl + 3",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true, Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }

        private void RegisterGlobalHotkey() { try { RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_3); } catch { } }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) { WakeUpWindow(); }
            base.WndProc(ref m);
        }
        private void WakeUpWindow() { if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal; this.Show(); this.Activate(); this.BringToFront(); }
        protected override void OnFormClosing(FormClosingEventArgs e) { try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { } base.OnFormClosing(e); }
    }
}
