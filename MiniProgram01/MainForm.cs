using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenu trayMenu;
    private TabControl tabControl;
    private string appName = "MiniProgram01";
    private static Color BgColor = Color.FromArgb(245, 245, 247);
    private HashSet<int> alertTabs = new HashSet<int>();
    private Timer flashTimer;
    private bool flashState = false;

    // --- 快捷鍵相關 API 與常數 ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int HOTKEY_ID = 9000;         // Ctrl+1: 喚醒視窗
    private const int HOTKEY_IMAGE_ID = 9001;   // Ctrl+2: 圖片小工具
    private const int HOTKEY_SAFETY_ID = 9002;  // Ctrl+3: Safety System
    private const int HOTKEY_GTASK_ID = 9008;   // Ctrl+9: G-Task (原 Ctrl+2)

    private const uint MOD_CONTROL = 0x0002;
    private const uint VK_1 = 0x31; 
    private const uint VK_2 = 0x32; 
    private const uint VK_3 = 0x33;
    private const uint VK_9 = 0x39;
    private const int WM_HOTKEY = 0x0312;

    // 用於儲存捷徑路徑的設定
    private Dictionary<int, string> hotkeyPaths = new Dictionary<int, string>();
    private string pathConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "hotkey_paths_config.txt");

    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
        // 初始化與載入路徑設定
        LoadPathSettings();

        this.Text = "整合通知中心";
        this.Width = 520; 
        this.Height = 560; 
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MinimumSize = new Size(450, 500); 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false; 
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Left = area.Right - this.Width - 10;
        this.Top = area.Bottom - this.Height - 10;

        // --- 設定常駐右鍵選單 (依照 Ctrl 順序排序) ---
        trayMenu = new ContextMenu();
        MenuItem startupItem = new MenuItem("開機自動啟動", ToggleStartup);
        startupItem.Checked = IsRunOnStartup();
        trayMenu.MenuItems.Add(startupItem);
        trayMenu.MenuItems.Add("-");
        
        trayMenu.MenuItems.Add("顯示主視窗 (Ctrl+1)", (s, e) => ShowAppWindow());
        trayMenu.MenuItems.Add("圖片小工具 (Ctrl+2)", (s, e) => LaunchExternalApp(HOTKEY_IMAGE_ID));
        trayMenu.MenuItems.Add("啟動 Safety System (Ctrl+3)", (s, e) => LaunchExternalApp(HOTKEY_SAFETY_ID));
        trayMenu.MenuItems.Add("啟動 G-Task (Ctrl+9)", (s, e) => LaunchExternalApp(HOTKEY_GTASK_ID));
        
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add("快捷鍵程式設定", (s, e) => OpenPathSettingsWindow());
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add("完全退出", (s, e) => { trayIcon.Visible = false; Environment.Exit(0); });
        
        trayIcon = new NotifyIcon();
        trayIcon.Icon = SystemIcons.Application; 
        trayIcon.ContextMenu = trayMenu;
        trayIcon.Visible = true;
        trayIcon.Text = "整合通知中心";
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();

        // --- 註冊全域快捷鍵 ---
        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);
        RegisterHotKey(this.Handle, HOTKEY_IMAGE_ID, MOD_CONTROL, VK_2);
        RegisterHotKey(this.Handle, HOTKEY_SAFETY_ID, MOD_CONTROL, VK_3); 
        RegisterHotKey(this.Handle, HOTKEY_GTASK_ID, MOD_CONTROL, VK_9); 

        // --- 初始化 TabControl ---
        tabControl = new TabControl();
        tabControl.Dock = DockStyle.Fill;
        tabControl.Font = new Font("Microsoft JhengHei UI", 10f);
        tabControl.ItemSize = new Size(80, 32);
        tabControl.Padding = new Point(0, 0);
        tabControl.SizeMode = TabSizeMode.Fixed; 
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += (s, e) => { 
            alertTabs.Remove(tabControl.SelectedIndex);
            tabControl.Invalidate(); 
        };
        this.Controls.Add(tabControl);

        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "轉待規");
        planApp = new App_TodoList(this, "plan", "轉待辦");
        todoApp.TargetList = planApp;
        planApp.TargetList = todoApp;
        recurringApp = new App_RecurringTasks(this, todoApp); 
        shortcutsApp = new App_Shortcuts(this);
        screenshotApp = new App_Screenshot(this);

        fileWatcherApp.Dock = DockStyle.Fill;
        todoApp.Dock = DockStyle.Fill;
        planApp.Dock = DockStyle.Fill;
        recurringApp.Dock = DockStyle.Fill;
        shortcutsApp.Dock = DockStyle.Fill;
        screenshotApp.Dock = DockStyle.Fill;

        tabControl.TabPages.Add(new TabPage("監控") { BackColor = BgColor });
        tabControl.TabPages.Add(new TabPage("待辦") { BackColor = BgColor });
        tabControl.TabPages.Add(new TabPage("待規") { BackColor = BgColor }); 
        tabControl.TabPages.Add(new TabPage("週期") { BackColor = BgColor });
        tabControl.TabPages.Add(new TabPage("捷徑") { BackColor = BgColor });
        tabControl.TabPages.Add(new TabPage("截圖") { BackColor = BgColor });

        tabControl.TabPages[0].Controls.Add(fileWatcherApp);
        tabControl.TabPages[1].Controls.Add(todoApp);
        tabControl.TabPages[2].Controls.Add(planApp); 
        tabControl.TabPages[3].Controls.Add(recurringApp);
        tabControl.TabPages[4].Controls.Add(shortcutsApp);
        tabControl.TabPages[5].Controls.Add(screenshotApp);

        flashTimer = new Timer() { Interval = 500 };
        flashTimer.Tick += (s, e) => {
            if (alertTabs.Count == 0) { 
                flashTimer.Stop();
                flashState = false; 
            } else { 
                flashState = !flashState;
                tabControl.Invalidate(); 
            }
        };
    }

    // 載入路徑設定
    private void LoadPathSettings() {
        // 預設路徑設定
        hotkeyPaths[HOTKEY_IMAGE_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MiniImage", "MiniImageStudio.exe");
        hotkeyPaths[HOTKEY_GTASK_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GTask", "GTaskNexus.exe");
        hotkeyPaths[HOTKEY_SAFETY_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetySystem", "Safety_System.exe");

        if (File.Exists(pathConfigFile)) {
            foreach (var line in File.ReadAllLines(pathConfigFile)) {
                var parts = line.Split('=');
                if (parts.Length == 2 && int.TryParse(parts[0], out int id)) {
                    hotkeyPaths[id] = parts[1];
                }
            }
        }
    }

    // 儲存路徑設定
    public void SavePathSettings(Dictionary<int, string> newPaths) {
        hotkeyPaths = newPaths;
        List<string> lines = new List<string>();
        foreach (var kvp in hotkeyPaths) {
            lines.Add($"{kvp.Key}={kvp.Value}");
        }
        File.WriteAllLines(pathConfigFile, lines);
    }
