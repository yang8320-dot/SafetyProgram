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

    // --- 原本的 Tab 繪製與提醒閃爍邏輯 ---
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        TabPage page = tabControl.TabPages[e.Index];
        bool isSelected = e.Index == tabControl.SelectedIndex;
        bool isAlert = alertTabs.Contains(e.Index);

        Color backColor = isSelected ? Color.White : BgColor;
        Color textColor = isSelected ? Color.Black : Color.Gray;

        if (isAlert && flashState && !isSelected) {
            backColor = Color.FromArgb(255, 200, 200); 
            textColor = Color.Red;
        }

        using (SolidBrush bgBrush = new SolidBrush(backColor)) { e.Graphics.FillRectangle(bgBrush, e.Bounds); }
        using (SolidBrush textBrush = new SolidBrush(textColor)) {
            StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            e.Graphics.DrawString(page.Text, e.Font, textBrush, e.Bounds, sf);
        }
    }

    // 修正的方法名稱：AlertTab (供其他模組呼叫)
    public void AlertTab(int tabIndex) {
        if (tabControl.SelectedIndex != tabIndex) {
            alertTabs.Add(tabIndex);
            if (!flashTimer.Enabled) flashTimer.Start();
        }
    }

    public void ClearTabAlert(int tabIndex) {
        alertTabs.Remove(tabIndex);
        tabControl.Invalidate();
    }

    // --- 視窗顯示與隱藏 ---
    // 修正的方法參數：支援傳入 tabIndex 自動切換分頁 (預設為 -1 表示不切換)
    public void ShowAppWindow(int tabIndex = -1) {
        if (tabIndex >= 0 && tabIndex < tabControl.TabPages.Count) {
            tabControl.SelectedIndex = tabIndex;
        }
        this.Show(); 
        this.WindowState = FormWindowState.Normal; 
        this.Activate(); 
    }

    private void HideAppWindow() { 
        this.Hide(); 
    }

    // --- 開機自動啟動邏輯 ---
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

    // --- 攔截全域熱鍵事件 ---
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID) { 
                // Ctrl+1: 喚醒主視窗
                ShowAppWindow();
            } else { 
                // 其他 ID (包含 Ctrl+2, Ctrl+3, Ctrl+9)
                LaunchExternalApp(id);
            }
        }
        base.WndProc(ref m);
    }

    // --- 執行外部程式邏輯 ---
    private void LaunchExternalApp(int hotkeyId) {
        if (hotkeyPaths.TryGetValue(hotkeyId, out string path) && !string.IsNullOrWhiteSpace(path)) {
            try {
                if (File.Exists(path)) {
                    Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true });
                } else {
                    MessageBox.Show($"找不到程式：\n{path}\n\n請確認路徑是否正確。", "執行失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } catch (Exception ex) {
                MessageBox.Show($"啟動程式時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } else {
            MessageBox.Show("尚未設定此快捷鍵的程式路徑！\n請對常駐圖示點擊右鍵 ->「快捷鍵程式設定」中設定。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // --- 開啟路徑設定視窗 ---
    private void OpenPathSettingsWindow() {
        using (var form = new HotkeyPathSettingsForm(hotkeyPaths, this)) {
            form.ShowDialog();
        }
    }

    // --- 攔截視窗關閉與縮小事件 ---
    protected override void OnFormClosing(FormClosingEventArgs e) {
        if (e.CloseReason == CloseReason.UserClosing) { 
            e.Cancel = true; 
            HideAppWindow(); // 點擊 X 時只隱藏，不關閉
        } 
        base.OnFormClosing(e);
    }

    protected override void OnResize(EventArgs e) {
        if (this.WindowState == FormWindowState.Minimized) { HideAppWindow(); } 
        base.OnResize(e);
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            // 程式結束前釋放熱鍵資源
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_IMAGE_ID);
            UnregisterHotKey(this.Handle, HOTKEY_SAFETY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_GTASK_ID);
            if (trayIcon != null) { trayIcon.Visible = false; trayIcon.Dispose(); }
            if (flashTimer != null) { flashTimer.Dispose(); }
        }
        base.Dispose(disposing);
    }
}

// =======================================================
// 新增的設定視窗類別 (HotkeyPathSettingsForm)
// =======================================================
public class HotkeyPathSettingsForm : Form {
    private Dictionary<int, string> currentPaths;
    private Dictionary<int, TextBox> textboxes = new Dictionary<int, TextBox>();
    private MainForm parentForm;

    public HotkeyPathSettingsForm(Dictionary<int, string> paths, MainForm parent) {
        this.currentPaths = new Dictionary<int, string>(paths);
        this.parentForm = parent;

        this.Text = "快捷鍵程式路徑設定";
        this.Width = 550;
        this.Height = 350;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;

        InitializeUI();
    }

    private void InitializeUI() {
        FlowLayoutPanel panel = new FlowLayoutPanel() {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            Padding = new Padding(20),
            AutoScroll = true
        };

        // 建立各快捷鍵的輸入區塊 (對應 ID)
        AddPathSettingUI(panel, 9001, "Ctrl + 2 (圖片小工具):");
        AddPathSettingUI(panel, 9002, "Ctrl + 3 (Safety System):");
        AddPathSettingUI(panel, 9008, "Ctrl + 9 (G-Task):");

        // 底部按鈕區塊
        Panel bottomPanel = new Panel() { Dock = DockStyle.Bottom, Height = 60 };
        Button btnSave = new Button() { Text = "儲存設定", Width = 100, Height = 35, Left = 160, Top = 10 };
        Button btnCancel = new Button() { Text = "取消", Width = 100, Height = 35, Left = 270, Top = 10 };

        btnSave.Click += BtnSave_Click;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        bottomPanel.Controls.Add(btnSave);
        bottomPanel.Controls.Add(btnCancel);

        this.Controls.Add(panel);
        this.Controls.Add(bottomPanel);
    }

    private void AddPathSettingUI(FlowLayoutPanel parent, int id, string labelText) {
        Panel row = new Panel() { Width = 480, Height = 55, Margin = new Padding(0, 5, 0, 5) };
        
        Label lbl = new Label() { Text = labelText, AutoSize = true, Location = new Point(5, 5), Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) };
        
        TextBox txt = new TextBox() { 
            Width = 380, 
            Location = new Point(5, 25), 
            Text = currentPaths.ContainsKey(id) ? currentPaths[id] : "" 
        };
        
        // 瀏覽檔案按鈕
        Button btnBrowse = new Button() { Text = "瀏覽...", Width = 70, Height = 25, Location = new Point(390, 24) };
        btnBrowse.Click += (s, e) => {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    txt.Text = ofd.FileName;
                }
            }
        };

        textboxes[id] = txt;
        
        row.Controls.Add(lbl);
        row.Controls.Add(txt);
        row.Controls.Add(btnBrowse);
        parent.Controls.Add(row);
    }

    private void BtnSave_Click(object sender, EventArgs e) {
        // 將 TextBox 的值寫回 Dictionary
        foreach (var kvp in textboxes) {
            currentPaths[kvp.Key] = kvp.Value.Text;
        }
        // 呼叫 MainForm 的儲存方法
        parentForm.SavePathSettings(currentPaths);
        MessageBox.Show("快捷鍵路徑設定已儲存！\n下次啟動時將自動載入。", "設定成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
