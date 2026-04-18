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
    // --- iOS 視覺語言定義 ---
    private static Color iosBackground = Color.FromArgb(242, 242, 247); // 系統底色
    private static Color iosCardWhite = Color.White;                    // 卡片底色
    private static Color iosAppleBlue = Color.FromArgb(0, 122, 255);    // 經典蘋果藍
    private static Color iosGray = Color.FromArgb(142, 142, 147);       // 輔助灰色
    private static Color iosRed = Color.FromArgb(255, 59, 48);          // 警告紅色
    private static Font iosFont = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);
    private static Font iosHeaderFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    public NotifyIcon trayIcon;
    public ContextMenu trayMenu;
    private TabControl tabControl;
    private string appName = "MiniProgram01";
    private HashSet<int> alertTabs = new HashSet<int>();
    private Timer flashTimer;
    private bool flashState = false;

    [cite_start]// --- 快捷鍵相關 API 與常數 [cite: 485-492] ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int HOTKEY_ID = 9000;         // Ctrl+1: 喚醒視窗
    private const int HOTKEY_IMAGE_ID = 9001;   // Ctrl+2: 圖片小工具
    private const int HOTKEY_SAFETY_ID = 9002;  // Ctrl+3: Safety System
    private const int HOTKEY_GTASK_ID = 9008;   // Ctrl+9: G-Task

    private const uint MOD_CONTROL = 0x0002;
    private const uint VK_1 = 0x31;
    private const uint VK_2 = 0x32; 
    private const uint VK_3 = 0x33;
    private const uint VK_9 = 0x39;
    private const int WM_HOTKEY = 0x0312;

    private Dictionary<int, string> hotkeyPaths = new Dictionary<int, string>();
    private string pathConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "hotkey_paths_config.txt");

    [cite_start]// --- 各個子模組 [cite: 493-494] ---
    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
        LoadPathSettings();
        
        [cite_start]// 視窗基礎設定[span_8](end_span)
        this.Text = "整合通知中心";
        this.Width = 540; 
        this.Height = 600; 
        this.BackColor = iosBackground;
        this.Font = iosFont;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MinimumSize = new Size(480, 520); 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false; 
        this.DoubleBuffered = true;
        this.AutoScaleMode = AutoScaleMode.Dpi; // 核心：支援 DPI 縮放

        [cite_start]// 定位至右下角 [cite: 496-497]
        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Left = area.Right - this.Width - 15;
        this.Top = area.Bottom - this.Height - 15;

        [cite_start]// 初始化右鍵選單 [cite: 498-501]
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

        [cite_start]// 註冊熱鍵 [cite: 503-504]
        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);
        RegisterHotKey(this.Handle, HOTKEY_IMAGE_ID, MOD_CONTROL, VK_2);
        RegisterHotKey(this.Handle, HOTKEY_SAFETY_ID, MOD_CONTROL, VK_3); 
        RegisterHotKey(this.Handle, HOTKEY_GTASK_ID, MOD_CONTROL, VK_9); 

        [cite_start]// 初始化 TabControl 並套用 iOS 扁平化樣式 [cite: 505-507]
        tabControl = new TabControl();
        tabControl.Dock = DockStyle.Fill;
        tabControl.Font = iosHeaderFont;
        tabControl.ItemSize = new Size(88, 42); // 加大觸控區域
        tabControl.SizeMode = TabSizeMode.Fixed; 
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed; // 自訂繪製標籤
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += (s, e) => { 
            alertTabs.Remove(tabControl.SelectedIndex);
            tabControl.Invalidate(); 
        };
        this.Controls.Add(tabControl);

        [cite_start]// 初始化子功能模組 [cite: 508-512]
        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "轉待規");
        planApp = new App_TodoList(this, "plan", "轉待辦");
        todoApp.TargetList = planApp;
        planApp.TargetList = todoApp;
        recurringApp = new App_RecurringTasks(this, todoApp); 
        shortcutsApp = new App_Shortcuts(this);
        screenshotApp = new App_Screenshot(this);

        AddTabPage("監控", fileWatcherApp);
        AddTabPage("待辦", todoApp);
        AddTabPage("待規", planApp); 
        AddTabPage("週期", recurringApp);
        AddTabPage("捷徑", shortcutsApp);
        AddTabPage("截圖", screenshotApp);

        [cite_start]// 閃爍提醒邏輯 [cite: 513-515]
        flashTimer = new Timer() { Interval = 500 };
        flashTimer.Tick += (s, e) => {
            if (alertTabs.Count == 0) { flashTimer.Stop(); flashState = false; }
            else { flashState = !flashState; tabControl.Invalidate(); }
        };
    }

    private void AddTabPage(string title, UserControl control) {
        TabPage page = new TabPage(title) { BackColor = iosBackground };
        control.Dock = DockStyle.Fill;
        page.Controls.Add(control);
        tabControl.TabPages.Add(page);
    }

    [cite_start]// 自定義繪製 iOS 風格分段控制標籤 [cite: 524-528]
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        Graphics g = e.Graphics;
        TabPage page = tabControl.TabPages[e.Index];
        Rectangle rect = tabControl.GetTabRect(e.Index);

        bool isSelected = (e.State == DrawItemState.Selected);
        bool isAlert = alertTabs.Contains(e.Index);

        Color backColor = isSelected ? iosCardWhite : iosBackground;
        Color textColor = isSelected ? iosAppleBlue : iosGray;

        if (isAlert && flashState && !isSelected) {
            backColor = Color.FromArgb(255, 235, 235);
            textColor = iosRed;
        }

        using (SolidBrush brush = new SolidBrush(backColor)) g.FillRectangle(brush, rect);

        TextRenderer.DrawText(g, page.Text, tabControl.Font, rect, textColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

        if (isSelected) {
            using (Pen p = new Pen(iosAppleBlue, 3)) {
                g.DrawLine(p, rect.Left + 10, rect.Bottom - 3, rect.Right - 10, rect.Bottom - 3);
            }
        }
    }

    [cite_start]// [cite: 529-532]
    public void AlertTab(int tabIndex) {
        if (tabControl.SelectedIndex != tabIndex) {
            alertTabs.Add(tabIndex);
            if (!flashTimer.Enabled) flashTimer.Start();
        }
    }

    public void ShowAppWindow(int tabIndex = -1) {
        if (tabIndex >= 0 && tabIndex < tabControl.TabPages.Count) tabControl.SelectedIndex = tabIndex;
        this.Show(); 
        this.WindowState = FormWindowState.Normal; 
        this.Activate();
    }

    private void LoadPathSettings() {
        hotkeyPaths[HOTKEY_IMAGE_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MiniImage", "MiniImageStudio.exe");
        hotkeyPaths[HOTKEY_GTASK_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GTask", "GTaskNexus.exe");
        hotkeyPaths[HOTKEY_SAFETY_ID] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetySystem", "Safety_System.exe");
        if (File.Exists(pathConfigFile)) {
            foreach (var line in File.ReadAllLines(pathConfigFile)) {
                var parts = line.Split('=');
                if (parts.Length == 2 && int.TryParse(parts[0], out int id)) hotkeyPaths[id] = parts[1];
            }
        }
    }

    public void SavePathSettings(Dictionary<int, string> newPaths) {
        hotkeyPaths = newPaths;
        File.WriteAllLines(pathConfigFile, hotkeyPaths.Select(kvp => $"{kvp.Key}={kvp.Value}"));
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID) ShowAppWindow();
            else LaunchExternalApp(id);
        }
        base.WndProc(ref m);
    }

    private void LaunchExternalApp(int hotkeyId) {
        if (hotkeyPaths.TryGetValue(hotkeyId, out string path) && File.Exists(path)) {
            try { Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show($"啟動錯誤：{ex.Message}"); }
        } else {
            MessageBox.Show("路徑未設定或檔案不存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = sender as MenuItem;
        bool newState = !item.Checked;
        SetRunOnStartup(newState);
        item.Checked = newState;
    }

    private bool IsRunOnStartup() {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false)) {
            return key?.GetValue(appName) != null;
        }
    }

    private void SetRunOnStartup(bool enable) {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)) {
            if (enable) key.SetValue(appName, Application.ExecutablePath);
            else key.DeleteValue(appName, false);
        }
    }

    private void OpenPathSettingsWindow() {
        using (var form = new HotkeyPathSettingsForm(hotkeyPaths, this)) form.ShowDialog();
    }

    protected override void OnFormClosing(FormClosingEventArgs e) {
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); }
        base.OnFormClosing(e);
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_IMAGE_ID);
            UnregisterHotKey(this.Handle, HOTKEY_SAFETY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_GTASK_ID);
            trayIcon?.Dispose();
            flashTimer?.Dispose();
        }
        base.Dispose(disposing);
    }
}

[cite_start]// 快捷鍵設定視窗 [cite: 554-573]
public class HotkeyPathSettingsForm : Form {
    private Dictionary<int, string> currentPaths;
    private Dictionary<int, TextBox> textboxes = new Dictionary<int, TextBox>();
    private MainForm parentForm;

    public HotkeyPathSettingsForm(Dictionary<int, string> paths, MainForm parent) {
        this.currentPaths = new Dictionary<int, string>(paths);
        this.parentForm = parent;
        this.Text = "快捷鍵路徑設定";
        this.Width = 500; this.Height = 320;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterParent;
        InitializeUI();
    }

    private void InitializeUI() {
        FlowLayoutPanel panel = new FlowLayoutPanel() { Dock = DockStyle.Fill, Padding = new Padding(25), FlowDirection = FlowDirection.TopDown };
        AddRow(panel, 9001, "Ctrl + 2 (圖片小工具):");
        AddRow(panel, 9002, "Ctrl + 3 (Safety System):");
        AddRow(panel, 9008, "Ctrl + 9 (G-Task):");

        Button btnSave = new Button() { Text = "儲存設定", Width = 120, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnSave.Click += (s, e) => {
            foreach (var kvp in textboxes) currentPaths[kvp.Key] = kvp.Value.Text;
            parentForm.SavePathSettings(currentPaths);
            this.Close();
        };
        panel.Controls.Add(btnSave);
        this.Controls.Add(panel);
    }

    private void AddRow(FlowLayoutPanel p, int id, string label) {
        p.Controls.Add(new Label() { Text = label, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) });
        TextBox txt = new TextBox() { Width = 380, Text = currentPaths.ContainsKey(id) ? currentPaths[id] : "" };
        textboxes[id] = txt;
        p.Controls.Add(txt);
        p.Controls.Add(new Label() { Height = 10 }); 
    }
}
