/*
 * 檔案功能：主視窗與整合通知中心 (提供系統列圖示、全域快捷鍵、分頁管理與 iOS 風格導覽列)
 * 對應選單名稱：主選單
 * 對應資料庫名稱：MainDB.sqlite (目前主框架以設定檔/記憶體為主)
 * 資料表名稱：MainForm_Config
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;

public class MainForm : Form
{
    // --- 介面核心控制項 ---
    private Panel topNavBar;                 // 模擬 iOS 風格的頂部導覽列
    private TabControl tabControl;           // 負責裝載各模組的容器 (隱藏原生標籤)
    private NotifyIcon trayIcon;             // 系統列常駐圖示
    private ContextMenu trayMenu;            // 右鍵選單
    private List<Button> navButtons = new List<Button>(); // 導覽列按鈕集合

    // --- 子模組參考 ---
    private UserControl fileWatcherApp;
    private dynamic todoApp;      // 使用 dynamic 方便未來的強型別綁定，或替換為實際類別名
    private dynamic planApp;
    private UserControl recurringApp;
    private UserControl shortcutsApp;
    private UserControl screenshotApp;

    // --- 狀態與常數 ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);
    private static Font ActiveFont = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold);

    private HashSet<int> alertTabs = new HashSet<int>();
    private System.Windows.Forms.Timer flashTimer; // 明確指定 Forms.Timer
    private bool flashState = false;

    // --- 全域快捷鍵 API ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    
    // 快捷鍵 ID 定義
    private const int HOTKEY_ID = 9000;       // Ctrl+1: 喚醒主視窗
    private const int HOTKEY_IMAGE_ID = 9001; // Ctrl+2
    private const int HOTKEY_SAFETY_ID = 9002;// Ctrl+3
    private const int HOTKEY_GTASK_ID = 9008; // Ctrl+9

    public MainForm()
    {
        // 1. 初始化表單基本設定與 DPI 支援
        this.Text = "整合通知中心";
        this.Width = 900;
        this.Height = 650;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = AppleBgColor;
        this.AutoScaleMode = AutoScaleMode.Dpi; // 【重要】DPI 縮放防模糊
        this.Font = MainFont;
        
        // 2. 初始化系統列圖示 (Tray Icon)
        InitializeTrayIcon();

        // 3. 初始化 iOS 風格介面與分頁容器
        InitializeUI();

        // 4. 註冊全域快捷鍵 (Ctrl = 2, 1 = 0x31)
        RegisterHotKey(this.Handle, HOTKEY_ID, 2, 0x31);
        RegisterHotKey(this.Handle, HOTKEY_IMAGE_ID, 2, 0x32);
        RegisterHotKey(this.Handle, HOTKEY_SAFETY_ID, 2, 0x33);
        RegisterHotKey(this.Handle, HOTKEY_GTASK_ID, 2, 0x39);

        // 5. 初始化閃爍提醒 Timer
        flashTimer = new System.Windows.Forms.Timer();
        flashTimer.Interval = 500; // 每 0.5 秒閃爍一次
        flashTimer.Tick += FlashTimer_Tick;
        flashTimer.Start();

        // 隱藏視窗作為背景常駐服務
        this.Load += (s, e) => { BeginInvoke(new Action(() => this.Hide())); };
    }

    /// <summary>
    /// 建構純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // 建立模擬 iOS Segmented Control 的導覽列
        topNavBar = new FlowLayoutPanel()
        {
            Dock = DockStyle.Top,
            Height = 55,
            BackColor = Color.White,
            Padding = new Padding(15, 10, 15, 0), // 內部與文字間隔
            Margin = new Padding(0, 0, 0, 15)     // 【關鍵】主選單與下方頁面間隔 15
        };
        
        // 底部加上一條極細的灰色分隔線提升立體感
        Panel borderLine = new Panel() { Dock = DockStyle.Top, Height = 1, BackColor = Color.FromArgb(230, 230, 230) };

        // 建立裝載各子模組的容器
        tabControl = new TabControl()
        {
            Dock = DockStyle.Fill,
            // 隱藏原生的醜陋分頁標籤，完全由上方的 topNavBar 控制
            Appearance = TabAppearance.FlatButtons,
            ItemSize = new Size(0, 1),
            SizeMode = TabSizeMode.Fixed,
            BackColor = AppleBgColor
        };

        // 解除警報事件：當切換分頁時，移除該分頁的警報狀態
        tabControl.SelectedIndexChanged += (s, e) => 
        {
            if (tabControl.SelectedIndex >= 0)
            {
                alertTabs.Remove(tabControl.SelectedIndex);
                SyncNavButtons(); // 同步更新按鈕的視覺狀態
            }
        };

        // 將控制項加入視窗主體
        this.Controls.Add(tabControl);
        this.Controls.Add(borderLine);
        this.Controls.Add(topNavBar);

        // 初始化各個獨立模組
        // (假設原本專案中的類別皆已存在，此處實例化並傳入必要的參考)
        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "轉待規");
        planApp = new App_TodoList(this, "plan", "轉待辦");
        todoApp.TargetList = planApp;
        planApp.TargetList = todoApp;
        recurringApp = new App_RecurringTasks(this, todoApp);
        shortcutsApp = new App_Shortcuts(this);
        screenshotApp = new App_Screenshot(this);

        // 設定各模組佈滿分頁
        fileWatcherApp.Dock = DockStyle.Fill;
        todoApp.Dock = DockStyle.Fill;
        planApp.Dock = DockStyle.Fill;
        recurringApp.Dock = DockStyle.Fill;
        shortcutsApp.Dock = DockStyle.Fill;
        screenshotApp.Dock = DockStyle.Fill;

        // 建立分頁並與按鈕綁定
        AddTab("檔案監控", fileWatcherApp);
        AddTab("待辦事項", todoApp);
        AddTab("計畫大綱", planApp);
        AddTab("週期任務", recurringApp);
        AddTab("常用捷徑", shortcutsApp);
        AddTab("畫面截圖", screenshotApp);

        // 預設選中第一個分頁
        if (navButtons.Count > 0) SyncNavButtons();
    }

    /// <summary>
    /// 動態新增分頁與對應的 iOS 風格導覽按鈕
    /// </summary>
    private void AddTab(string title, Control content)
    {
        // 建立分頁
        TabPage page = new TabPage(title) { BackColor = AppleBgColor };
        page.Controls.Add(content);
        tabControl.TabPages.Add(page);

        int tabIndex = tabControl.TabPages.Count - 1;

        // 建立導覽列按鈕
        Button btn = new Button()
        {
            Text = title,
            AutoSize = true,
            Height = 35,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Padding = new Padding(10, 0, 10, 0), // 【關鍵】框內與文字間隔 10
            Margin = new Padding(0, 0, 10, 0),
            Font = MainFont,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100) // 預設未選中顏色
        };
        btn.FlatAppearance.BorderSize = 0; // 隱藏邊框

        // 點擊按鈕時切換 Tab
        btn.Click += (s, e) => 
        {
            tabControl.SelectedIndex = tabIndex;
            SyncNavButtons();
        };

        navButtons.Add(btn);
        topNavBar.Controls.Add(btn);
    }

    /// <summary>
    /// 同步導覽列按鈕狀態 (高亮顯示目前所在的分頁)
    /// </summary>
    private void SyncNavButtons()
    {
        for (int i = 0; i < navButtons.Count; i++)
        {
            if (i == tabControl.SelectedIndex)
            {
                // 啟動狀態：Apple 經典藍色，粗體
                navButtons[i].ForeColor = AppleBlue;
                navButtons[i].Font = ActiveFont;
            }
            else
            {
                // 若該分頁正處於警報狀態且正在閃爍，保留警報色
                if (alertTabs.Contains(i) && flashState)
                {
                    navButtons[i].ForeColor = Color.IndianRed;
                }
                else
                {
                    // 恢復未選中顏色
                    navButtons[i].ForeColor = Color.FromArgb(100, 100, 100);
                    navButtons[i].Font = MainFont;
                }
            }
        }
    }

    /// <summary>
    /// 初始化系統列右鍵選單與圖示
    /// </summary>
    private void InitializeTrayIcon()
    {
        trayMenu = new ContextMenu();
        
        MenuItem mnuShow = new MenuItem("開啟控制面板", (s, e) => ShowAppWindow());
        MenuItem mnuStartup = new MenuItem("開機自動啟動", ToggleStartup);
        mnuStartup.Checked = IsRunOnStartup(); // 檢查登錄檔狀態
        MenuItem mnuExit = new MenuItem("完全退出程式", (s, e) => Application.Exit());

        trayMenu.MenuItems.Add(mnuShow);
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add(mnuStartup);
        trayMenu.MenuItems.Add("-");
        trayMenu.MenuItems.Add(mnuExit);

        trayIcon = new NotifyIcon();
        trayIcon.Text = "整合通知中心 (執行中)";
        trayIcon.Icon = SystemIcons.Application; // 可自行替換為專案資源 .ico
        trayIcon.ContextMenu = trayMenu;
        trayIcon.Visible = true;
        
        // 左鍵雙擊喚醒主視窗
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();
    }

    // --- 核心邏輯操作 ---

    /// <summary>
    /// 顯示並喚醒主視窗 (執行緒安全)
    /// </summary>
    public void ShowAppWindow(int targetTabIndex = -1)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => ShowAppWindow(targetTabIndex)));
            return;
        }

        if (targetTabIndex >= 0 && targetTabIndex < tabControl.TabPages.Count)
        {
            tabControl.SelectedIndex = targetTabIndex;
            SyncNavButtons();
        }

        this.Show();
        this.WindowState = FormWindowState.Normal;
        this.Activate(); // 將視窗推至最前端
    }

    /// <summary>
    /// 處理閃爍邏輯 (執行緒安全)
    /// </summary>
    private void FlashTimer_Tick(object sender, EventArgs e)
    {
        if (alertTabs.Count == 0) return;

        flashState = !flashState;
        SyncNavButtons(); // 觸發按鈕重新渲染顏色
    }

    /// <summary>
    /// 外部模組呼叫：為特定分頁加入警報閃爍狀態
    /// </summary>
    public void AddAlertTab(int tabIndex)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => AddAlertTab(tabIndex)));
            return;
        }

        if (tabIndex >= 0 && tabIndex < tabControl.TabPages.Count)
        {
            alertTabs.Add(tabIndex);
        }
    }

    // --- 開機啟動註冊邏輯 ---
    private void ToggleStartup(object sender, EventArgs e)
    {
        MenuItem item = sender as MenuItem;
        bool newState = !item.Checked;
        SetRunOnStartup(newState);
        item.Checked = newState;
    }

    private bool IsRunOnStartup()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false))
        {
            if (key != null)
            {
                return key.GetValue("MiniProgram_NotifyCenter") != null;
            }
        }
        return false;
    }

    private void SetRunOnStartup(bool enable)
    {
        string path = Application.ExecutablePath;
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
        {
            if (enable) key.SetValue("MiniProgram_NotifyCenter", "\"" + path + "\"");
            else key.DeleteValue("MiniProgram_NotifyCenter", false);
        }
    }

    // --- 覆寫作業系統熱鍵訊息攔截 ---
    protected override void WndProc(ref Message m)
    {
        const int WM_HOTKEY = 0x0312;
        if (m.Msg == WM_HOTKEY)
        {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID)
            {
                ShowAppWindow();
            }
            // 於此處可依據專案需求補充其他快捷鍵觸發邏輯，如呼叫圖片截取等
        }
        base.WndProc(ref m);
    }

    // --- 表單生命週期管理 ---
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // 點擊右上角關閉時，僅隱藏視窗而非關閉程式
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            this.Hide();
        }
        base.OnFormClosing(e);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // 釋放資源與快捷鍵註冊
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_IMAGE_ID);
            UnregisterHotKey(this.Handle, HOTKEY_SAFETY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_GTASK_ID);

            if (trayIcon != null) trayIcon.Dispose();
            if (flashTimer != null) flashTimer.Dispose();
        }
        base.Dispose(disposing);
    }
}
