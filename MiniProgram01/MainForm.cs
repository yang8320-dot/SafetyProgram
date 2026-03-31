using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenu trayMenu;
    private TabControl tabControl;
    private string appName = "MiniProgram01";
    private static Color BgColor = Color.FromArgb(245, 245, 247); 

    private HashSet<int> alertTabs = new HashSet<int>();
    private Timer flashTimer;
    private bool flashState = false;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int HOTKEY_ID = 9000;
    private const uint MOD_CONTROL = 0x0002; 
    private const uint VK_1 = 0x31; 
    private const int WM_HOTKEY = 0x0312;

    public MainForm() {
        this.Text = "整合通知中心";
        this.Width = 380; this.Height = 520;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true; this.ShowInTaskbar = false; 
        this.BackColor = BgColor;

        // 只有「剛啟動時」預設放在右下角，之後由使用者自由決定位置
        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);

        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "整合通知中心" };
        trayMenu = new ContextMenu();
        trayMenu.MenuItems.Add("顯示主面板", new EventHandler(delegate { ShowAppWindow(); }));
        trayMenu.MenuItems.Add("-");

        tabControl = new TabControl() { 
            Dock = DockStyle.Fill, Padding = new Point(12, 5), 
            Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold),
            DrawMode = TabDrawMode.OwnerDrawFixed
        };
        tabControl.DrawItem += TabControl_DrawItem;
        
        tabControl.SelectedIndexChanged += (s, e) => {
            if (alertTabs.Contains(tabControl.SelectedIndex)) {
                alertTabs.Remove(tabControl.SelectedIndex);
                tabControl.Invalidate(tabControl.GetTabRect(tabControl.SelectedIndex));
                tabControl.Update();
            }
        };
        this.Controls.Add(tabControl);

        flashTimer = new Timer() { Interval = 500, Enabled = true };
        flashTimer.Tick += (s, e) => {
            if (alertTabs.Count > 0) {
                flashState = !flashState;
                foreach (int index in alertTabs) {
                    tabControl.Invalidate(tabControl.GetTabRect(index));
                }
                tabControl.Update(); 
            }
        };

        // 1. 監控系統
        TabPage tabWatcher = new TabPage("監控系統");
        App_FileWatcher watcherApp = new App_FileWatcher(this, trayMenu);
        watcherApp.Dock = DockStyle.Fill;
        tabWatcher.Controls.Add(watcherApp);
        tabControl.TabPages.Add(tabWatcher);

        // 2. 待辦清單
        TabPage tabTodo = new TabPage("待辦清單");
        App_TodoList todoApp = new App_TodoList(this);
        todoApp.Dock = DockStyle.Fill;
        tabTodo.Controls.Add(todoApp);
        tabControl.TabPages.Add(tabTodo);

        // 3. 週期任務
        TabPage tabRecurring = new TabPage("週期任務");
        App_RecurringTasks recurringApp = new App_RecurringTasks(this, todoApp);
        recurringApp.Dock = DockStyle.Fill;
        tabRecurring.Controls.Add(recurringApp);
        tabControl.TabPages.Add(tabRecurring);

        // 4. 捷徑管理
        TabPage tabShortcuts = new TabPage("捷徑管理");
        App_Shortcuts shortcutsApp = new App_Shortcuts(this);
        shortcutsApp.Dock = DockStyle.Fill;
        tabShortcuts.Controls.Add(shortcutsApp);
        tabControl.TabPages.Add(tabShortcuts);

        // 5. 截圖功能 (全新加入)
        TabPage tabScreenshot = new TabPage("截圖");
        App_Screenshot screenshotApp = new App_Screenshot(this);
        screenshotApp.Dock = DockStyle.Fill;
        tabScreenshot.Controls.Add(screenshotApp);
        tabControl.TabPages.Add(tabScreenshot);

        // 右下角常駐選單設定
        trayMenu.MenuItems.Add("-");
        MenuItem startupMenu = new MenuItem("開機自動執行") { Checked = IsRunOnStartup() };
        startupMenu.Click += new EventHandler(ToggleStartup);
        trayMenu.MenuItems.Add(startupMenu);
        trayMenu.MenuItems.Add("結束程式", new EventHandler(delegate { UnregisterHotKey(this.Handle, HOTKEY_ID); trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = trayMenu;

        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { this.Hide(); this.Opacity = 1; });
    }

    public void AlertTab(int tabIndex) {
        if (tabControl.SelectedIndex != tabIndex && tabIndex >= 0 && tabIndex < tabControl.TabCount) {
            alertTabs.Add(tabIndex);
            flashState = true;
            tabControl.Invalidate(tabControl.GetTabRect(tabIndex));
            tabControl.Update();
        }
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        Graphics g = e.Graphics;
        TabPage page = tabControl.TabPages[e.Index];
        Rectangle bounds = tabControl.GetTabRect(e.Index);

        bool isAlert = alertTabs.Contains(e.Index) && flashState;
        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

        Color bgColor = isAlert ? Color.Crimson : (isSelected ? Color.White : Color.FromArgb(240, 240, 240));
        Color textColor = isAlert ? Color.White : (isSelected ? Color.FromArgb(0, 122, 255) : Color.Black);

        using (Brush b = new SolidBrush(bgColor)) { g.FillRectangle(b, bounds); }
        if (isSelected) {
            using (Pen p = new Pen(Color.FromArgb(0, 122, 255), 3)) {
                g.DrawLine(p, bounds.Left, bounds.Top, bounds.Right, bounds.Top);
            }
        }
        TextRenderer.DrawText(g, page.Text, page.Font, bounds, textColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }

    public void ShowAppWindow(int targetTabIndex = -1) {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        
        // 自動切換到指定的分頁 (用於截圖完畢後自動跳回截圖頁)
        if (targetTabIndex >= 0 && targetTabIndex < tabControl.TabCount) {
            tabControl.SelectedIndex = targetTabIndex; 
        }
        
        this.Activate(); 
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) {
            this.Hide(); 
            this.WindowState = FormWindowState.Normal; 
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) { ShowAppWindow(); }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e) { 
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } 
        base.OnFormClosing(e); 
    }

    private bool IsRunOnStartup() {
        try { using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) { 
            if (rk != null && rk.GetValue(appName) != null) return true;
        } } catch { } return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender; item.Checked = !item.Checked;
        try { using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) { 
            if (item.Checked) rk.SetValue(appName, "\"" + Application.ExecutablePath + "\""); 
            else if (rk != null) rk.DeleteValue(appName, false); 
        } } catch {}
    }
}
