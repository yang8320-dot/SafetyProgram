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

    // 閃動提醒相關變數
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

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);

        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "整合通知中心" };
        trayMenu = new ContextMenu();
        trayMenu.MenuItems.Add("顯示主面板", new EventHandler(delegate { ShowAppWindow(); }));
        trayMenu.MenuItems.Add("-");

        // 開啟自行繪製 Tab 模式
        tabControl = new TabControl() { 
            Dock = DockStyle.Fill, Padding = new Point(12, 5), 
            Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold),
            DrawMode = TabDrawMode.OwnerDrawFixed // 允許接管標籤繪圖
        };
        tabControl.DrawItem += TabControl_DrawItem;
        
        // 切換分頁時，解除該分頁的閃爍狀態並強制重繪該標籤
        tabControl.SelectedIndexChanged += (s, e) => {
            if (alertTabs.Contains(tabControl.SelectedIndex)) {
                alertTabs.Remove(tabControl.SelectedIndex);
                tabControl.Invalidate(tabControl.GetTabRect(tabControl.SelectedIndex));
                tabControl.Update();
            }
        };
        this.Controls.Add(tabControl);

        // 【修正核心】閃爍計時器：精準鎖定並強制重繪特定的「標籤區塊」
        flashTimer = new Timer() { Interval = 500, Enabled = true };
        flashTimer.Tick += (s, e) => {
            if (alertTabs.Count > 0) {
                flashState = !flashState;
                foreach (int index in alertTabs) {
                    // 只針對需要閃爍的標籤區塊發出重繪要求
                    tabControl.Invalidate(tabControl.GetTabRect(index));
                }
                tabControl.Update(); // 強制系統立即將顏色畫上去
            }
        };

        // 載入四大模組
        TabPage tabWatcher = new TabPage("📁 監控");
        App_FileWatcher watcherApp = new App_FileWatcher(this, trayMenu);
        watcherApp.Dock = DockStyle.Fill;
        tabWatcher.Controls.Add(watcherApp);
        tabControl.TabPages.Add(tabWatcher);

        TabPage tabTodo = new TabPage("📝 待辦");
        App_TodoList todoApp = new App_TodoList(this);
        todoApp.Dock = DockStyle.Fill;
        tabTodo.Controls.Add(todoApp);
        tabControl.TabPages.Add(tabTodo);

        TabPage tabRecurring = new TabPage("🔁 週期");
        App_RecurringTasks recurringApp = new App_RecurringTasks(this, todoApp);
        recurringApp.Dock = DockStyle.Fill;
        tabRecurring.Controls.Add(recurringApp);
        tabControl.TabPages.Add(tabRecurring);

        TabPage tabShortcuts = new TabPage("🚀 捷徑");
        App_Shortcuts shortcutsApp = new App_Shortcuts(this);
        shortcutsApp.Dock = DockStyle.Fill;
        tabShortcuts.Controls.Add(shortcutsApp);
        tabControl.TabPages.Add(tabShortcuts);

        trayMenu.MenuItems.Add("-");
        MenuItem startupMenu = new MenuItem("開機自動執行") { Checked = IsRunOnStartup() };
        startupMenu.Click += new EventHandler(ToggleStartup);
        trayMenu.MenuItems.Add(startupMenu);
        trayMenu.MenuItems.Add("結束程式", new EventHandler(delegate { UnregisterHotKey(this.Handle, HOTKEY_ID); trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = trayMenu;

        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { this.Hide(); this.Opacity = 1; });
    }

    // 呼叫此方法可讓指定的分頁標籤進入閃爍狀態
    public void AlertTab(int tabIndex) {
        if (tabControl.SelectedIndex != tabIndex && tabIndex >= 0 && tabIndex < tabControl.TabCount) {
            alertTabs.Add(tabIndex);
            flashState = true;
            // 收到通知瞬間，立刻強制畫上警示色
            tabControl.Invalidate(tabControl.GetTabRect(tabIndex));
            tabControl.Update();
        }
    }

    // 【修正核心】自繪分頁邏輯：強化閃爍色彩的對比度
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        Graphics g = e.Graphics;
        TabPage page = tabControl.TabPages[e.Index];
        Rectangle bounds = tabControl.GetTabRect(e.Index);

        bool isAlert = alertTabs.Contains(e.Index) && flashState;
        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

        Color bgColor;
        Color textColor;

        // 判斷目前標籤應該長什麼樣子
        if (isAlert) {
            bgColor = Color.Crimson; // 顯眼的深紅色
            textColor = Color.White; // 白色文字
        } else if (isSelected) {
            bgColor = Color.White;
            textColor = Color.FromArgb(0, 122, 255); // 蘋果藍
        } else {
            bgColor = Color.FromArgb(240, 240, 240); // 預設灰
            textColor = Color.Black;
        }

        // 1. 填滿背景色
        using (Brush b = new SolidBrush(bgColor)) { 
            g.FillRectangle(b, bounds); 
        }

        // 2. 如果是被選取狀態，畫一條藍色頂線增加設計感
        if (isSelected) {
            using (Pen p = new Pen(Color.FromArgb(0, 122, 255), 3)) {
                g.DrawLine(p, bounds.Left, bounds.Top, bounds.Right, bounds.Top);
            }
        }

        // 3. 畫上文字
        TextRenderer.DrawText(g, page.Text, page.Font, bounds, textColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }

    public void ShowAppWindow(int targetTabIndex = -1) {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        if (targetTabIndex >= 0 && targetTabIndex < tabControl.TabCount) tabControl.SelectedIndex = targetTabIndex; 
        this.Activate(); 
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
