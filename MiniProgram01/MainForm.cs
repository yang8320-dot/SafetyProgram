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

    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
        this.Text = "整合通知中心";
        // 【修正1】將主視窗稍微加寬至 460，確保容納 6 個分頁與內容
        this.Width = 460; this.Height = 520;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true; this.ShowInTaskbar = false; 
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Left = area.Right - this.Width - 10; 
        this.Top = area.Bottom - this.Height - 10;

        trayMenu = new ContextMenu();
        MenuItem startupItem = new MenuItem("開機自動啟動", ToggleStartup);
        startupItem.Checked = IsRunOnStartup();
        trayMenu.MenuItems.Add(startupItem);
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add("顯示主視窗", (s, e) => ShowAppWindow());
        trayMenu.MenuItems.Add("完全退出", (s, e) => { trayIcon.Visible = false; Environment.Exit(0); });
        
        trayIcon = new NotifyIcon() { Icon = SystemIcons.Application, ContextMenu = trayMenu, Visible = true, Text = "整合通知中心 (Ctrl+1)" };
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();

        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);

        // 【修正2】強制鎖定 Tab 寬度，平均分配 6 個，徹底消滅滾動箭頭
        tabControl = new TabControl() { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 10f), ItemSize = new Size(70, 30), Padding = new Point(0, 0) };
        tabControl.SizeMode = TabSizeMode.Fixed; 
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += (s, e) => { alertTabs.Remove(tabControl.SelectedIndex); tabControl.Invalidate(); };
        this.Controls.Add(tabControl);

        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "轉待規");
        planApp = new App_TodoList(this, "plan", "轉待辦");
        todoApp.TargetList = planApp;
        planApp.TargetList = todoApp;
        recurringApp = new App_RecurringTasks(this, todoApp); 
        shortcutsApp = new App_Shortcuts(this);
        screenshotApp = new App_Screenshot(this);

        // 【修正3】強制所有模組填滿畫面，解決破圖與重疊問題
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
            if (alertTabs.Count == 0) { flashTimer.Stop(); flashState = false; }
            else { flashState = !flashState; tabControl.Invalidate(); }
        };
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

    public void ShowAppWindow(int targetTabIndex = -1) {
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        if (!this.Visible) this.Show();
        if (targetTabIndex >= 0 && targetTabIndex < tabControl.TabCount) { tabControl.SelectedIndex = targetTabIndex; }
        this.Activate(); 
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) { this.Hide(); this.WindowState = FormWindowState.Normal; }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) { ShowAppWindow(); }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
