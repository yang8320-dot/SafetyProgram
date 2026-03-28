using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenu trayMenu;
    private TabControl tabControl;
    private bool isPositionLocked = false;
    private string appName = "MiniProgram01";
    private static Color BgColor = Color.FromArgb(245, 245, 247); 

    public MainForm() {
        IntPtr forceHandle = this.Handle;
        this.Text = "整合通知中心";
        this.Width = 380; // 稍微拉寬一點容納頁籤
        this.Height = 450;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MaximizeBox = true;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        // --- 建立全域右鍵選單 ---
        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "整合通知中心" };
        trayMenu = new ContextMenu();
        
        trayMenu.MenuItems.Add("顯示主面板", new EventHandler(delegate { ShowAppWindow(); }));
        trayMenu.MenuItems.Add("-");

        // --- 建立頁籤系統 ---
        tabControl = new TabControl() { Dock = DockStyle.Fill, Padding = new Point(15, 6), Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) };
        this.Controls.Add(tabControl);

        // 【模組 1】載入檔案監控
        TabPage tabWatcher = new TabPage("📁 檔案監控");
        tabWatcher.BackColor = BgColor;
        App_FileWatcher watcherApp = new App_FileWatcher(this, trayMenu);
        watcherApp.Dock = DockStyle.Fill;
        tabWatcher.Controls.Add(watcherApp);
        tabControl.TabPages.Add(tabWatcher);

        // 【模組 2】載入待辦事項
        TabPage tabTodo = new TabPage("📝 待辦事項");
        tabTodo.BackColor = BgColor;
        App_TodoList todoApp = new App_TodoList(this);
        todoApp.Dock = DockStyle.Fill;
        tabTodo.Controls.Add(todoApp);
        tabControl.TabPages.Add(tabTodo);

        // --- 全域設定選單 ---
        trayMenu.MenuItems.Add("-");
        MenuItem startupMenu = new MenuItem("開機自動執行");
        startupMenu.Checked = IsRunOnStartup(); 
        startupMenu.Click += new EventHandler(ToggleStartup);
        trayMenu.MenuItems.Add(startupMenu);

        MenuItem lockMenu = new MenuItem("鎖定視窗位置");
        lockMenu.Click += new EventHandler(delegate { isPositionLocked = !isPositionLocked; lockMenu.Checked = isPositionLocked; });
        trayMenu.MenuItems.Add(lockMenu);
        
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add("結束程式", new EventHandler(delegate { trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = trayMenu;

        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { this.Hide(); this.Opacity = 1; });
    }

    public void ShowAppWindow() {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        this.Activate(); 
        this.Refresh();  
    }

    protected override void WndProc(ref Message m) {
        const int WM_NCLBUTTONDOWN = 0x00A1; const int HTCAPTION = 2; const int WM_SYSCOMMAND = 0x0112; const int SC_MOVE = 0xF010;          
        if (isPositionLocked) {
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt32() & 0xfff0) == SC_MOVE) return;
            if (m.Msg == WM_NCLBUTTONDOWN && m.WParam.ToInt32() == HTCAPTION) return;
        }
        base.WndProc(ref m);
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) { this.Hide(); this.WindowState = FormWindowState.Normal; }
    }

    protected override void OnFormClosing(FormClosingEventArgs e) { 
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } 
        base.OnFormClosing(e); 
    }

    private bool IsRunOnStartup() {
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) {
                if (rk != null && rk.GetValue(appName) != null) { return rk.GetValue(appName).ToString().Replace("\"", "").Equals(Application.ExecutablePath, StringComparison.OrdinalIgnoreCase); }
            }
        } catch {} return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender; item.Checked = !item.Checked;
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) {
                if (item.Checked) { rk.SetValue(appName, "\"" + Application.ExecutablePath + "\""); trayIcon.ShowBalloonTip(3000, "設定成功", "已設定開機執行！", ToolTipIcon.Info); } 
                else { rk.DeleteValue(appName, false); trayIcon.ShowBalloonTip(3000, "設定成功", "已取消開機執行。", ToolTipIcon.Info); }
            }
        } catch {}
    }
}
