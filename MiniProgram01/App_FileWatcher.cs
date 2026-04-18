// ============================================================
// FILE: MiniProgram01/App_FileWatcher.cs 
// ============================================================
using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private ContextMenu trayMenu;
    private MenuItem trayToggleWatcherItem;

    private FileSystemWatcher watcher;
    private string watchPath = "";
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "watcher_config.txt");

    // --- iOS 風格色彩與字體定義 ---
    private static Color iosBackground = Color.FromArgb(242, 242, 247);
    private static Color iosCardWhite = Color.White;
    private static Color iosAppleBlue = Color.FromArgb(0, 122, 255);
    private static Color iosGreen = Color.FromArgb(52, 199, 89);
    private static Color iosRed = Color.FromArgb(255, 59, 48);
    private static Color iosText = Color.FromArgb(28, 28, 30);
    private static Color iosGrayText = Color.FromArgb(142, 142, 147);

    private static Font MainFont = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    // --- 介面元件 ---
    private TextBox txtPath;
    private Button btnSelectFolder;
    private Button btnToggle;
    private Button btnClearLog;
    private ListBox logBox;
    private Label lblStatus;

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.trayMenu = menu;

        // 核心：支援高 DPI 縮放
        this.AutoScaleMode = AutoScaleMode.Dpi; 
        this.BackColor = iosBackground;
        this.Padding = new Padding(15);
        this.Font = MainFont;

        InitializeUI();
        LoadConfig();
        InitializeTrayMenu();

        if (!string.IsNullOrEmpty(watchPath) && Directory.Exists(watchPath)) {
            StartWatching();
        } else {
            UpdateStatus(false);
        }
    }

    private void InitializeUI() {
        // --- 頂部控制面板 (卡片風格) ---
        Panel topCard = new Panel() {
            Dock = DockStyle.Top,
            Height = 130,
            BackColor = iosCardWhite,
            Padding = new Padding(15),
            Margin = new Padding(0, 0, 0, 15)
        };

        // 標題與狀態
        Label lblTitle = new Label() {
            Text = "資料夾監控",
            Font = new Font("Microsoft JhengHei UI", 14f, FontStyle.Bold),
            ForeColor = iosText,
            AutoSize = true,
            Location = new Point(15, 15)
        };
        topCard.Controls.Add(lblTitle);

        lblStatus = new Label() {
            Text = "狀態：未啟動",
            Font = BoldFont,
            ForeColor = iosRed,
            AutoSize = true,
            Location = new Point(130, 18)
        };
        topCard.Controls.Add(lblStatus);

        // 路徑輸入框與瀏覽按鈕
        txtPath = new TextBox() {
            Location = new Point(15, 55),
            Width = 320,
            Height = 35,
            Font = MainFont,
            BorderStyle = BorderStyle.FixedSingle,
            ReadOnly = true,
            BackColor = iosBackground
        };
        topCard.Controls.Add(txtPath);

        btnSelectFolder = CreateIOSButton("選擇目錄", iosGrayText, Color.FromArgb(230, 230, 235));
        btnSelectFolder.Location = new Point(345, 54);
        btnSelectFolder.Width = 80;
        btnSelectFolder.Click += (s, e) => {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                fbd.Description = "請選擇要監控的資料夾";
                if (fbd.ShowDialog() == DialogResult.OK) {
                    watchPath = fbd.SelectedPath;
                    txtPath.Text = watchPath;
                    SaveConfig();
                    if (watcher != null && watcher.EnableRaisingEvents) StartWatching(); 
                }
            }
        };
        topCard.Controls.Add(btnSelectFolder);

        // 啟動/停止按鈕
        btnToggle = CreateIOSButton("啟動監控", Color.White, iosAppleBlue);
        btnToggle.Location = new Point(15, 95);
        btnToggle.Width = 120;
        btnToggle.Click += (s, e) => {
            if (watcher != null && watcher.EnableRaisingEvents) StopWatching();
            else StartWatching();
        };
        topCard.Controls.Add(btnToggle);

        // 清除日誌按鈕
        btnClearLog = CreateIOSButton("清除紀錄", Color.White, iosGrayText);
        btnClearLog.Location = new Point(145, 95);
        btnClearLog.Width = 100;
        btnClearLog.Click += (s, e) => { logBox.Items.Clear(); };
        topCard.Controls.Add(btnClearLog);

        this.Controls.Add(topCard);

        // --- 中間間距 ---
        Panel spacer = new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = iosBackground };
        this.Controls.Add(spacer);

        // --- 下方日誌顯示區 (卡片風格) ---
        Panel bottomCard = new Panel() {
            Dock = DockStyle.Fill,
            BackColor = iosCardWhite,
            Padding = new Padding(15)
        };

        Label lblLogTitle = new Label() {
            Text = "監控日誌",
            Font = BoldFont,
            ForeColor = iosText,
            Dock = DockStyle.Top,
            Height = 30
        };
        bottomCard.Controls.Add(lblLogTitle);

        logBox = new ListBox() {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = iosCardWhite,
            ForeColor = iosText,
            Font = MainFont,
            ItemHeight = 22,
            HorizontalScrollbar = true
        };
        
        // 雙擊開啟路徑功能
        logBox.DoubleClick += (s, e) => {
            if (logBox.SelectedItem != null) {
                string itemText = logBox.SelectedItem.ToString();
                int pathIndex = itemText.IndexOf("] ") + 2;
                if (pathIndex > 1 && pathIndex < itemText.Length) {
                    string filePath = itemText.Substring(pathIndex).Split(' ')[0]; // 嘗試截取路徑
                    if (File.Exists(filePath) || Directory.Exists(filePath)) {
                        try { Process.Start("explorer.exe", $"/select,\"{filePath}\""); } catch { }
                    }
                }
            }
        };

        bottomCard.Controls.Add(logBox);
        this.Controls.Add(bottomCard);
    }

    private Button CreateIOSButton(string text, Color foreColor, Color backColor) {
        Button btn = new Button() {
            Text = text,
            Height = 32,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = backColor,
            ForeColor = foreColor,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold)
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // --- 系統列右鍵選單整合 ---
    private void InitializeTrayMenu() {
        if (trayMenu == null) return;
        trayMenu.MenuItems.Add("-");
        trayToggleWatcherItem = new MenuItem("啟動/停止檔案監控", (s, e) => {
            if (watcher != null && watcher.EnableRaisingEvents) StopWatching();
            else StartWatching();
        });
        trayMenu.MenuItems.Add(trayToggleWatcherItem);
    }

    // --- 監控核心邏輯 ---
    private void StartWatching() {
        if (string.IsNullOrEmpty(watchPath) || !Directory.Exists(watchPath)) {
            MessageBox.Show("請先選擇有效的資料夾路徑！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            UpdateStatus(false);
            return;
        }

        if (watcher == null) {
            watcher = new FileSystemWatcher();
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite;
            watcher.Created += OnFileChanged;
            watcher.Deleted += OnFileChanged;
            watcher.Changed += OnFileChanged;
            watcher.Renamed += OnFileRenamed;
        }

        watcher.Path = watchPath;
        watcher.EnableRaisingEvents = true;
        UpdateStatus(true);
        LogMessage($"[系統] 開始監控：{watchPath}");
    }

    private void StopWatching() {
        if (watcher != null) {
            watcher.EnableRaisingEvents = false;
        }
        UpdateStatus(false);
        LogMessage("[系統] 監控已停止。");
    }

    private void UpdateStatus(bool isRunning) {
        if (this.InvokeRequired) {
            this.Invoke(new Action(() => UpdateStatus(isRunning)));
            return;
        }

        if (isRunning) {
            lblStatus.Text = "狀態：監控中";
            lblStatus.ForeColor = iosGreen;
            btnToggle.Text = "停止監控";
            btnToggle.BackColor = iosRed;
            if (trayToggleWatcherItem != null) trayToggleWatcherItem.Checked = true;
        } else {
            lblStatus.Text = "狀態：未啟動";
            lblStatus.ForeColor = iosRed;
            btnToggle.Text = "啟動監控";
            btnToggle.BackColor = iosAppleBlue;
            if (trayToggleWatcherItem != null) trayToggleWatcherItem.Checked = false;
        }
    }

    // --- 檔案事件處理 (包含跨執行緒 UI 更新與 Alert) ---
    private void OnFileChanged(object sender, FileSystemEventArgs e) {
        string action = e.ChangeType == WatcherChangeTypes.Created ? "新增" :
                        e.ChangeType == WatcherChangeTypes.Deleted ? "刪除" : "修改";
        
        LogMessage($"[{action}] {e.FullPath}");
        
        // 如果是新增檔案，觸發 MainForm 標籤閃爍警告 (假設監控在 Index 0)
        if (e.ChangeType == WatcherChangeTypes.Created) {
            if (parentForm != null) {
                parentForm.Invoke(new Action(() => parentForm.AlertTab(0)));
            }
        }
    }

    private void OnFileRenamed(object sender, RenamedEventArgs e) {
        LogMessage($"[重新命名] {e.OldFullPath} -> {e.FullPath}");
    }

    private void LogMessage(string msg) {
        if (logBox.InvokeRequired) {
            logBox.Invoke(new Action(() => LogMessage(msg)));
            return;
        }

        string timeStr = DateTime.Now.ToString("HH:mm:ss");
        logBox.Items.Insert(0, $"{timeStr} {msg}");
        
        // 限制日誌筆數避免記憶體爆增
        if (logBox.Items.Count > 1000) {
            logBox.Items.RemoveAt(logBox.Items.Count - 1);
        }
    }

    // --- 存檔與載入設定 ---
    private void LoadConfig() {
        if (File.Exists(configFile)) {
            try {
                watchPath = File.ReadAllText(configFile).Trim();
                txtPath.Text = watchPath;
            } catch { }
        }
    }

    private void SaveConfig() {
        try {
            File.WriteAllText(configFile, watchPath);
        } catch { }
    }
}
