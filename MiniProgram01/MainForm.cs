using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenu trayMenu;
    private TabControl tabControl;
    private bool isPositionLocked = false;
    private string appName = "MiniProgram01";
    private static Color BgColor = Color.FromArgb(245, 245, 247); 

    // --- 快捷鍵相關 P/Invoke ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    private const int HOTKEY_ID = 9000;
    private const uint MOD_ALT = 0x0001;
    private const uint VK_1 = 0x31; // '1' 鍵
    private const int WM_HOTKEY = 0x0312;

    public MainForm() {
        this.Text = "整合通知中心";
        this.Width = 380; 
        this.Height = 500;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        // 註冊預設快捷鍵 Alt + 1
        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_ALT, VK_1);

        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "整合通知中心" };
        trayMenu = new ContextMenu();
        trayMenu.MenuItems.Add("顯示主面板", new EventHandler(delegate { ShowAppWindow(); }));
        trayMenu.MenuItems.Add("-");

        tabControl = new TabControl() { Dock = DockStyle.Fill, Padding = new Point(15, 6), Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) };
        this.Controls.Add(tabControl);

        // 【模組 1】檔案監控
        TabPage tabWatcher = new TabPage("📁 檔案監控");
        tabWatcher.Controls.Add(new App_FileWatcher(this, trayMenu) { Dock = DockStyle.Fill });
        tabControl.TabPages.Add(tabWatcher);

        // 【模組 2】待辦事項
        TabPage tabTodo = new TabPage("📝 待辦");
        App_TodoList todoApp = new App_TodoList(this);
        tabTodo.Controls.Add(todoApp { Dock = DockStyle.Fill });
        tabControl.TabPages.Add(tabTodo);

        // 【模組 3】週期任務
        TabPage tabRecurring = new TabPage("🔁 週期任務");
        tabRecurring.Controls.Add(new App_RecurringTasks(this, todoApp) { Dock = DockStyle.Fill });
        tabControl.TabPages.Add(tabRecurring);

        // 【全新模組 4】捷徑管理
        TabPage tabShortcuts = new TabPage("🚀 捷徑");
        tabShortcuts.Controls.Add(new App_Shortcuts(this) { Dock = DockStyle.Fill });
        tabControl.TabPages.Add(tabShortcuts);

        trayMenu.MenuItems.Add("-");
        MenuItem startupMenu = new MenuItem("開機自動執行") { Checked = IsRunOnStartup() };
        startupMenu.Click += new EventHandler(ToggleStartup);
        trayMenu.MenuItems.Add(startupMenu);
        trayMenu.MenuItems.Add("結束程式", new EventHandler(delegate { UnregisterHotKey(this.Handle, HOTKEY_ID); trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = trayMenu;

        this.Opacity = 0; 
        this.Load += (s, e) => { this.Hide(); this.Opacity = 1; };
    }

    public void ShowAppWindow(int targetTabIndex = -1) {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        if (targetTabIndex >= 0) tabControl.SelectedIndex = targetTabIndex; // 切換分頁
        this.Activate(); 
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) {
            ShowAppWindow(1); // 按下快捷鍵，呼叫程式並跳到「待辦(Index 1)」
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e) { 
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } 
        base.OnFormClosing(e); 
    }

    private bool IsRunOnStartup() {
        try { using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) { return rk?.GetValue(appName) != null; } } catch { return false; }
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender; item.Checked = !item.Checked;
        try { using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) { if (item.Checked) rk.SetValue(appName, "\"" + Application.ExecutablePath + "\""); else rk.DeleteValue(appName, false); } } catch {}
    }
}
