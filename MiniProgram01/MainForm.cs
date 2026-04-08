using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;

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

    private const int HOTKEY_ID = 9000;
    private const int HOTKEY_GTASK_ID = 9001; 
    private const int HOTKEY_SAFETY_ID = 9002; 

    private const uint MOD_CONTROL = 0x0002; 
    private const uint VK_1 = 0x31; 
    private const uint VK_2 = 0x32; 
    private const uint VK_3 = 0x33; 
    private const int WM_HOTKEY = 0x0312;

    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
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

        // --- 設定常駐右鍵選單 ---
        trayMenu = new ContextMenu();
        
        MenuItem startupItem = new MenuItem("開機自動啟動", ToggleStartup);
        startupItem.Checked = IsRunOnStartup();
        trayMenu.MenuItems.Add(startupItem);
        trayMenu.MenuItems.Add("-");
        
        trayMenu.MenuItems.Add("顯示主視窗 (Ctrl+1)", (s, e) => ShowAppWindow());
        trayMenu.MenuItems.Add("啟動 G-Task (Ctrl+2)", (s, e) => LaunchGTask());
        trayMenu.MenuItems.Add("啟動 Safety System (Ctrl+3)", (s, e) => LaunchSafetySystem()); 
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
        RegisterHotKey(this.Handle, HOTKEY_GTASK_ID, MOD_CONTROL, VK_2); 
        RegisterHotKey(this.Handle, HOTKEY_SAFETY_ID, MOD_CONTROL, VK_3); 

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

    private void LaunchGTask() {
        try {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GTask", "GTaskNexus.exe");
            if (File.Exists(path)) {
                Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true });
            } else {
                MessageBox.Show("找不到 G-Task 檔案：\n" + path, "啟動失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        } catch (Exception ex) {
            MessageBox.Show("錯誤: " + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LaunchSafetySystem() {
        try {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetySystem", "Safety_System.exe");
            if (File.Exists(path)) {
                Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true });
            } else {
                MessageBox.Show("找不到 Safety System 檔案：\n" + path, "啟動失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        } catch (Exception ex) {
            MessageBox.Show("錯誤: " + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        Graphics g = e.Graphics; TabPage page = tabControl.TabPages[e.Index]; Rectangle rect = e.Bounds;
        bool isSelected = (tabControl.SelectedIndex == e.Index);
        bool isAlert = alertTabs.Contains(e.Index) && flashState;
        Color bg = isAlert ? Color.IndianRed : (isSelected ? Color.White : Color.FromArgb(230, 230, 230));
        Color fg = isAlert ? Color.White : (isSelected ? Color.FromArgb(0, 122, 255) : Color.Gray);
        g.FillRectangle(new SolidBrush(bg), rect);
        if (isSelected && !isAlert) g.FillRectangle(new SolidBrush(Color.FromArgb(0, 122, 255)), rect.Left, rect.Bottom - 3, rect.Width, 3);
        StringFormat sf = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(page.Text, new Font(tabControl.Font, isSelected ? FontStyle.Bold : FontStyle.Regular), new SolidBrush(fg), rect, sf);
    }

    public void AlertTab(int targetTabIndex) {
        if (tabControl.SelectedIndex != targetTabIndex) {
            alertTabs.Add(targetTabIndex);
            if (!flashState) { flashState = true; flashTimer.Start(); }
            tabControl.Invalidate();
        }
    }

    // 【核心修正 1】：喚醒視窗邏輯，加入尺寸防呆
    public void ShowAppWindow(int targetTabIndex = -1) {
        if (!this.Visible) this.Show();
        
        if (this.WindowState == FormWindowState.Minimized) {
            this.WindowState = FormWindowState.Normal;
        }

        // 雙重防呆：如果 Windows 還是把視窗縮小到極限，強制拉回正常寬高
        if (this.Width < 450 || this.Height < 500) {
            this.Width = 520;
            this.Height = 560;
        }

        if (targetTabIndex >= 0 && targetTabIndex < tabControl.TabCount) {
            tabControl.SelectedIndex = targetTabIndex;
        }
        
        this.Activate(); 
        this.BringToFront(); // 確保跳到最上層
    }

    // 【核心修正 2】：最小化時的邏輯處理
    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) { 
            this.Hide(); 
            // 絕對不能在這裡把 WindowState 設回 Normal，會導致尺寸錯亂！
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            int hotkeyId = m.WParam.ToInt32();
            if (hotkeyId == HOTKEY_ID) ShowAppWindow(); 
            else if (hotkeyId == HOTKEY_GTASK_ID) LaunchGTask();
            else if (hotkeyId == HOTKEY_SAFETY_ID) LaunchSafetySystem();
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e) { 
        if (e.CloseReason == CloseReason.UserClosing) { 
            e.Cancel = true; 
            // 這裡如果是按 X 關閉，也只是隱藏視窗
            this.Hide(); 
        } 
        base.OnFormClosing(e); 
    }

    private bool IsRunOnStartup() {
        try { 
            RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false);
            if (rk != null) { bool exists = (rk.GetValue(appName) != null); rk.Close(); return exists; }
        } catch { } return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender; item.Checked = !item.Checked;
        try { 
            RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            if (rk != null) {
                if (item.Checked) rk.SetValue(appName, "\"" + Application.ExecutablePath + "\"");
                else rk.DeleteValue(appName, false);
                rk.Close();
            }
        } catch { MessageBox.Show("權限不足。"); }
    }
}
