/*
 * 檔案功能：系統主視窗框架 (最終完整版)
 * 包含模組：iOS 風格導覽、系統列常駐、全域快捷鍵設定、開機自動啟動註冊
 */

using System;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32; 

public class MainForm : Form
{
    // --- Windows API 宣告 (全域快捷鍵) ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int MOD_CONTROL = 0x0002;
    private const int HOTKEY_ID_AWAKE = 1;
    private const int HOTKEY_ID_APP2 = 2;
    private const int HOTKEY_ID_APP3 = 3;
    private const int HOTKEY_ID_APP9 = 9;

    // --- 快捷鍵目標路徑 ---
    private string pathApp2 = "";
    private string pathApp3 = "";
    private string pathApp9 = "";

    // --- 介面元件 ---
    private NotifyIcon trayIcon;
    private ContextMenuStrip trayMenu; // 【修正】改用現代化 ContextMenuStrip
    private Panel navBar;
    private Panel contentPanel;
    private App_TodoList appTodo;

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);

    public MainForm()
    {
        // 1. 視窗基本設定
        this.Text = "整合通知中心";
        this.Width = 900;
        this.Height = 700;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;

        // 2. 建立系統列 (Tray) 與開機啟動選單
        InitializeTrayIcon();

        // 3. 建立 UI 架構 (左側導覽、右側內容)
        InitializeUI();

        // 4. 註冊全域快捷鍵
        RegisterAllHotkeys();

        // 5. 載入資料庫中的快捷鍵設定
        LoadHotkeySettings();

        // 攔截視窗關閉事件，改為縮小至系統列
        this.FormClosing += (s, e) =>
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
        };
    }

    private void InitializeUI()
    {
        navBar = new Panel() 
        { 
            Dock = DockStyle.Left, 
            Width = 150, 
            BackColor = Color.White, 
            Padding = new Padding(10) 
        };
        
        contentPanel = new Panel() 
        { 
            Dock = DockStyle.Fill, 
            BackColor = AppleBgColor 
        };

        this.Controls.Add(contentPanel);
        this.Controls.Add(navBar);

        // 實例化各模組
        appTodo = new App_TodoList(this, "todo", "轉至計畫");
        var appPlan = new App_TodoList(this, "plan", "轉至待辦");
        var appShortcuts = new App_Shortcuts(this);
        var appWatcher = new App_FileWatcher(this, trayMenu);
        var appTasks = new App_RecurringTasks(this, appTodo);
        var appScreen = new App_Screenshot(this);

        // 將模組加入導覽列
        AddNavButton("待辦事項", appTodo);
        AddNavButton("計畫大綱", appPlan);
        AddNavButton("常用捷徑", appShortcuts);
        AddNavButton("檔案監控", appWatcher);
        AddNavButton("週期任務", appTasks);
        AddNavButton("畫面截圖", appScreen);

        // 預設選取第一個模組
        if (navBar.Controls.Count > 0 && navBar.Controls[navBar.Controls.Count - 1] is Button firstBtn)
        {
            firstBtn.PerformClick();
        }
    }

    private void AddNavButton(string title, UserControl module)
    {
        module.Dock = DockStyle.Fill;

        Button btn = new Button()
        {
            Text = title,
            Dock = DockStyle.Top,
            Height = 45,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = Color.White,
            ForeColor = Color.Black,
            Font = new Font(MainFont.FontFamily, 11f, FontStyle.Regular),
            Margin = new Padding(0, 0, 0, 10)
        };
        btn.FlatAppearance.BorderSize = 0;
        
        btn.Click += (s, e) =>
        {
            foreach (Control c in navBar.Controls)
            {
                if (c is Button b)
                {
                    b.BackColor = Color.White;
                    b.ForeColor = Color.Black;
                    b.Font = new Font(MainFont.FontFamily, 11f, FontStyle.Regular);
                }
            }
            
            btn.BackColor = AppleBlue;
            btn.ForeColor = Color.White;
            btn.Font = new Font(MainFont.FontFamily, 11f, FontStyle.Bold);

            contentPanel.Controls.Clear();
            contentPanel.Controls.Add(module);
        };

        navBar.Controls.Add(btn);
        btn.BringToFront(); 
    }

    // ==========================================
    // 系統列 (Tray) 與全域設定選單
    // ==========================================
    private void InitializeTrayIcon()
    {
        trayMenu = new ContextMenuStrip(); // 【修正】使用 ContextMenuStrip
        
        // 【修正】使用 Items.Add 並對應現代化的寫法
        trayMenu.Items.Add("開啟主畫面", null, (s, e) => ShowAppWindow());
        trayMenu.Items.Add("快捷鍵路徑設定", null, (s, e) => { new HotkeySettingsWindow(this).ShowDialog(); });
        
        // 開機自動啟動選項
        ToolStripMenuItem startupItem = new ToolStripMenuItem("開機自動啟動"); // 【修正】使用 ToolStripMenuItem
        startupItem.Checked = IsRunOnStartup();
        startupItem.Click += ToggleStartup;
        trayMenu.Items.Add(startupItem);

        trayMenu.Items.Add(new ToolStripSeparator()); // 【修正】加入分隔線
        
        trayMenu.Items.Add("完全退出程式", null, (s, e) => 
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_AWAKE);
            UnregisterHotKey(this.Handle, HOTKEY_ID_APP2);
            UnregisterHotKey(this.Handle, HOTKEY_ID_APP3);
            UnregisterHotKey(this.Handle, HOTKEY_ID_APP9);
            trayIcon.Visible = false;
            Application.Exit();
        });

        trayIcon = new NotifyIcon()
        {
            Text = "整合通知中心 (背景執行中)",
            Icon = SystemIcons.Application,
            ContextMenuStrip = trayMenu, // 【修正】將 ContextMenu 改為 ContextMenuStrip 屬性
            Visible = true
        };
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();
    }

    // --- 開機自動啟動 (Registry 登錄檔控制) ---
    private void ToggleStartup(object sender, EventArgs e)
    {
        ToolStripMenuItem item = sender as ToolStripMenuItem; // 【修正】轉型為 ToolStripMenuItem
        bool newState = !item.Checked;
        SetRunOnStartup(newState);
        item.Checked = newState;
    }

    private bool IsRunOnStartup()
    {
        try 
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false))
            {
                return key?.GetValue("整合通知中心") != null;
            }
        }
        catch { return false; }
    }

    private void SetRunOnStartup(bool enable)
    {
        try 
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
            {
                if (enable)
                    key.SetValue("整合通知中心", Application.ExecutablePath);
                else
                    key.DeleteValue("整合通知中心", false);
            }
        }
        catch (Exception ex) 
        { 
            MessageBox.Show($"設定開機啟動失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
        }
    }

    public void ShowAppWindow()
    {
        this.Show();
        this.WindowState = FormWindowState.Normal;
        this.Activate();
    }

    // ==========================================
    // 全域快捷鍵邏輯與 SQLite 存取
    // ==========================================
    private void RegisterAllHotkeys()
    {
        RegisterHotKey(this.Handle, HOTKEY_ID_AWAKE, MOD_CONTROL, (int)Keys.D1); 
        RegisterHotKey(this.Handle, HOTKEY_ID_APP2, MOD_CONTROL, (int)Keys.D2);  
        RegisterHotKey(this.Handle, HOTKEY_ID_APP3, MOD_CONTROL, (int)Keys.D3);  
        RegisterHotKey(this.Handle, HOTKEY_ID_APP9, MOD_CONTROL, (int)Keys.D9);  
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_HOTKEY = 0x0312;
        if (m.Msg == WM_HOTKEY)
        {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID_AWAKE) 
                ShowAppWindow();
            else if (id == HOTKEY_ID_APP2 && !string.IsNullOrEmpty(pathApp2)) 
                LaunchApp(pathApp2);
            else if (id == HOTKEY_ID_APP3 && !string.IsNullOrEmpty(pathApp3)) 
                LaunchApp(pathApp3);
            else if (id == HOTKEY_ID_APP9 && !string.IsNullOrEmpty(pathApp9)) 
                LaunchApp(pathApp9);
        }
        base.WndProc(ref m);
    }

    private void LaunchApp(string path)
    {
        try 
        { 
            Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true }); 
        }
        catch 
        { 
            MessageBox.Show($"無法啟動目標: {path}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
        }
    }

    public void LoadHotkeySettings()
    {
        try
        {
            using (var conn = DatabaseManager.GetConnection())
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT SettingKey, SettingValue FROM GlobalSettings WHERE SettingKey LIKE 'Hotkey%'", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string k = reader.GetString(0); 
                        string v = reader.GetString(1);
                        if (k == "Hotkey2_Path") pathApp2 = v;
                        else if (k == "Hotkey3_Path") pathApp3 = v;
                        else if (k == "Hotkey9_Path") pathApp9 = v;
                    }
                }
            }
        }
        catch { }
    }

    public async Task SaveHotkeySettingsAsync(string p2, string p3, string p9)
    {
        this.pathApp2 = p2; 
        this.pathApp3 = p3; 
        this.pathApp9 = p9;
        
        await Task.Run(() =>
        {
            using (var conn = DatabaseManager.GetConnection())
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    void Set(string k, string v)
                    {
                        using (var cmd = new SQLiteCommand("INSERT OR REPLACE INTO GlobalSettings (SettingKey, SettingValue) VALUES (@k, @v)", conn, tx))
                        { 
                            cmd.Parameters.AddWithValue("@k", k); 
                            cmd.Parameters.AddWithValue("@v", v); 
                            cmd.ExecuteNonQuery(); 
                        }
                    }
                    Set("Hotkey2_Path", p2); 
                    Set("Hotkey3_Path", p3); 
                    Set("Hotkey9_Path", p9);
                    tx.Commit();
                }
            }
        });
    }

    public void AddAlertTab(int tabIndex)
    {
        if (this.InvokeRequired) 
        { 
            this.Invoke(new Action(() => AddAlertTab(tabIndex))); 
            return; 
        }
        
        foreach (Control c in navBar.Controls)
        {
            if (c is Button b && b.Text == "檔案監控" && b.BackColor != AppleBlue)
            {
                b.BackColor = Color.IndianRed;
                b.ForeColor = Color.White;
            }
        }
    }
}

public class HotkeySettingsWindow : Form
{
    private MainForm parent;
    private TextBox txtH2, txtH3, txtH9;

    public HotkeySettingsWindow(MainForm parent)
    {
        this.parent = parent;
        this.Text = "全域快捷鍵路徑設定";
        this.Width = 500; 
        this.Height = 400; 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog; 
        this.MaximizeBox = false; 
        this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); 
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        parent.LoadHotkeySettings();

        FlowLayoutPanel flow = new FlowLayoutPanel() 
        { 
            Dock = DockStyle.Fill, 
            FlowDirection = FlowDirection.TopDown, 
            Padding = new Padding(20) 
        };
        
        flow.Controls.Add(new Label() 
        { 
            Text = "設定 Ctrl + 2 / 3 / 9 所對應啟動的程式或檔案路徑：", 
            AutoSize = true, 
            Margin = new Padding(0, 0, 0, 20), 
            ForeColor = Color.Gray 
        });

        txtH2 = AddRow(flow, "Ctrl + 2 (如圖片小工具)：", "Hotkey2_Path");
        txtH3 = AddRow(flow, "Ctrl + 3 (如 Safety System)：", "Hotkey3_Path");
        txtH9 = AddRow(flow, "Ctrl + 9 (如 G-Task)：", "Hotkey9_Path");

        Button btnSave = new Button() 
        { 
            Text = "儲存設定", 
            Width = 440, 
            Height = 45, 
            BackColor = Color.FromArgb(0, 122, 255), 
            ForeColor = Color.White, 
            FlatStyle = FlatStyle.Flat, 
            Cursor = Cursors.Hand, 
            Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), 
            Margin = new Padding(0, 20, 0, 0) 
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) => 
        { 
            await parent.SaveHotkeySettingsAsync(txtH2.Text, txtH3.Text, txtH9.Text); 
            MessageBox.Show("設定已成功儲存至資料庫！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
            this.Close(); 
        };
        
        flow.Controls.Add(btnSave); 
        this.Controls.Add(flow);
    }

    private TextBox AddRow(FlowLayoutPanel parentPanel, string label, string settingKey)
    {
        parentPanel.Controls.Add(new Label() 
        { 
            Text = label, 
            AutoSize = true, 
            Margin = new Padding(0, 0, 0, 5) 
        });
        
        TableLayoutPanel row = new TableLayoutPanel() 
        { 
            Width = 440, 
            Height = 35, 
            ColumnCount = 2, 
            Margin = new Padding(0, 0, 0, 15) 
        };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
        
        string val = "";
        using (var conn = DatabaseManager.GetConnection()) 
        {
            conn.Open();
            using (var cmd = new SQLiteCommand("SELECT SettingValue FROM GlobalSettings WHERE SettingKey = @k", conn)) 
            {
                cmd.Parameters.AddWithValue("@k", settingKey);
                var res = cmd.ExecuteScalar(); 
                if (res != null) val = res.ToString();
            }
        }

        TextBox tb = new TextBox() 
        { 
            Dock = DockStyle.Fill, 
            Text = val, 
            BorderStyle = BorderStyle.FixedSingle, 
            Margin = new Padding(0, 5, 10, 0) 
        };
        
        Button btn = new Button() 
        { 
            Text = "瀏覽", 
            Dock = DockStyle.Fill, 
            FlatStyle = FlatStyle.Flat, 
            Cursor = Cursors.Hand, 
            BackColor = Color.White 
        };
        btn.FlatAppearance.BorderColor = Color.Gray;
        btn.Click += (s, e) => 
        { 
            OpenFileDialog ofd = new OpenFileDialog(); 
            if (ofd.ShowDialog() == DialogResult.OK) tb.Text = ofd.FileName; 
        };
        
        row.Controls.Add(tb, 0, 0); 
        row.Controls.Add(btn, 1, 0); 
        parentPanel.Controls.Add(row); 
        
        return tb;
    }
}
